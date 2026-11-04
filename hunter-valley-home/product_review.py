#!/usr/bin/env python3
"""
Hunter Valley Home - supplier (Winwin) product review sheet.

Builds ./output/hunter_valley_home_product_review.xlsx (+ .csv) with one row per
ACTIVE product, so the supplier can mark which products ship to Australia via
normal channels.

READ-ONLY: this script only ever sends GraphQL *queries* to Shopify. It never
calls a mutation and never writes to the store.

Data sources (first one available wins, or force with --source):
  1. api  - Shopify Admin GraphQL API. Needs SHOPIFY_STORE + SHOPIFY_ADMIN_TOKEN
            in .env (or the environment). Token needs read_products scope.
  2. json - A JSON dump in the same shape the API path produces
            (written by --save-json, or by any other read-only export).
  3. csv  - Shopify product CSV export at ./input/products_export.csv.

Dependencies: openpyxl, requests (api source only). Standard library otherwise.
"""
from __future__ import annotations

import argparse
import csv
import html
import json
import os
import re
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path

STORE_DOMAIN = "huntervalleyhome.com"
DEFAULT_API_VERSION = "2026-07"
DEFAULT_INPUT_CSV = Path("input/products_export.csv")
DEFAULT_OUTPUT_DIR = Path("output")
OUTPUT_BASENAME = "hunter_valley_home_product_review"

OVERSIZE_CM = 60.0
OVERSIZE_KG = 10.0

COLUMNS = [
    "Product URL",
    "Product title",
    "Product type / category",
    "Material(s)",
    "Dimensions (L x W x H, cm)",
    "Weight (kg)",
    "Number of variants",
    "SKU(s)",
    "Selling price (AUD)",
    "Main image URL",
    "Contains wood?",
    "Oversized?",
    "Can ship normally to AU? (Winwin to fill)",
    "Reason if not (Winwin to fill)",
    "Notes (Winwin to fill)",
]
WINWIN_COLUMNS = COLUMNS[12:15]
INTERNAL_COLUMNS = [
    "product_id",
    "handle",
    "status",
    "collections",
    "title",
    "vendor",
    "wood_evidence",
    "material_source",
    "dimensions_source",
    "weight_source",
]


# --------------------------------------------------------------------------- #
# Normalised product model
# --------------------------------------------------------------------------- #
@dataclass
class Variant:
    sku: str = ""
    price: float | None = None
    weight_kg: float | None = None
    title: str = ""
    options: dict = field(default_factory=dict)


@dataclass
class Product:
    product_id: str = ""
    handle: str = ""
    title: str = ""
    status: str = ""
    product_type: str = ""
    category: str = ""  # Shopify standard taxonomy full name, if any
    vendor: str = ""
    tags: list = field(default_factory=list)
    body_html: str = ""
    image_url: str = ""
    collections: list = field(default_factory=list)
    metafields: dict = field(default_factory=dict)  # "ns.key" -> (type, value)
    variants: list = field(default_factory=list)
    variants_count: int | None = None


# --------------------------------------------------------------------------- #
# .env loading (no python-dotenv dependency)
# --------------------------------------------------------------------------- #
def load_dotenv(paths) -> None:
    for p in paths:
        p = Path(p)
        if not p.is_file():
            continue
        for line in p.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            key = key.strip().removeprefix("export ").strip()
            value = value.strip().strip('"').strip("'")
            os.environ.setdefault(key, value)


# --------------------------------------------------------------------------- #
# Source 1: Shopify Admin GraphQL API (read-only queries)
# --------------------------------------------------------------------------- #
PRODUCTS_QUERY = """
query Products($first: Int!, $cursor: String) {
  products(first: $first, after: $cursor, query: "status:active", sortKey: ID) {
    pageInfo { hasNextPage endCursor }
    nodes {
      id legacyResourceId handle title status productType tags descriptionHtml vendor
      category { fullName }
      featuredMedia { preview { image { url } } }
      collections(first: 15) { nodes { title handle } }
      metafields(first: 30) { nodes { namespace key type value } }
      variantsCount { count }
    }
  }
}
"""

VARIANTS_QUERY = """
query Variants($first: Int!, $cursor: String) {
  productVariants(first: $first, after: $cursor, sortKey: ID) {
    pageInfo { hasNextPage endCursor }
    nodes {
      sku price title
      selectedOptions { name value }
      product { id }
      inventoryItem { measurement { weight { unit value } } }
    }
  }
}
"""

SHOP_QUERY = "query { shop { name currencyCode primaryDomain { host } } }"


class ShopifyGraphQL:
    def __init__(self, store: str, token: str, api_version: str):
        import requests  # imported lazily so CSV mode needs no requests

        store = store.strip().removeprefix("https://").removeprefix("http://").rstrip("/")
        if not store.endswith(".myshopify.com"):
            store = f"{store}.myshopify.com"
        self.url = f"https://{store}/admin/api/{api_version}/graphql.json"
        self.session = requests.Session()
        self.session.headers.update(
            {"X-Shopify-Access-Token": token, "Content-Type": "application/json"}
        )

    def query(self, query: str, variables: dict | None = None) -> dict:
        # Safety net: this tool must stay read-only.
        if re.match(r"\s*mutation\b", query):
            raise RuntimeError("Refusing to send a mutation - this script is read-only.")
        for attempt in range(8):
            resp = self.session.post(
                self.url, json={"query": query, "variables": variables or {}}, timeout=60
            )
            if resp.status_code == 429 or resp.status_code >= 500:
                time.sleep(2 ** min(attempt, 4))
                continue
            if resp.status_code == 401:
                raise SystemExit("Shopify API: 401 Unauthorized - check SHOPIFY_ADMIN_TOKEN.")
            if resp.status_code == 403:
                raise SystemExit("Shopify API: 403 Forbidden - token needs the read_products scope.")
            resp.raise_for_status()
            payload = resp.json()
            errors = payload.get("errors") or []
            if any((e.get("extensions") or {}).get("code") == "THROTTLED" for e in errors):
                time.sleep(2 + attempt)
                continue
            if errors:
                raise RuntimeError(f"GraphQL errors: {json.dumps(errors)[:800]}")
            throttle = ((payload.get("extensions") or {}).get("cost") or {}).get("throttleStatus")
            if throttle and throttle.get("currentlyAvailable", 1000) < 200:
                time.sleep(max(1.0, (200 - throttle["currentlyAvailable"]) / throttle.get("restoreRate", 50)))
            return payload["data"]
        raise RuntimeError("Shopify API kept throttling / failing - giving up.")

    def paginate(self, query: str, root: str, page_size: int):
        cursor = None
        while True:
            try:
                data = self.query(query, {"first": page_size, "cursor": cursor})
            except RuntimeError as exc:
                if "MAX_COST_EXCEEDED" in str(exc) and page_size > 1:
                    page_size = max(1, page_size // 2)
                    continue
                raise
            conn = data[root]
            yield from conn["nodes"]
            if not conn["pageInfo"]["hasNextPage"]:
                break
            cursor = conn["pageInfo"]["endCursor"]


def fetch_from_api(store: str, token: str, api_version: str) -> dict:
    client = ShopifyGraphQL(store, token, api_version)
    shop = client.query(SHOP_QUERY)["shop"]
    print(f"Connected to {shop['name']} ({shop['primaryDomain']['host']}), currency {shop['currencyCode']}")
    products = []
    for node in client.paginate(PRODUCTS_QUERY, "products", 15):
        products.append(node)
        if len(products) % 75 == 0:
            print(f"  ...{len(products)} active products")
    print(f"Fetched {len(products)} active products; fetching variants...")
    active_ids = {p["id"] for p in products}
    variants = []
    for node in client.paginate(VARIANTS_QUERY, "productVariants", 150):
        if node["product"]["id"] in active_ids:
            variants.append(node)
    print(f"Fetched {len(variants)} variants of active products")
    return {"shop": shop, "products": products, "variants": variants}


WEIGHT_TO_KG = {
    "GRAMS": 0.001, "G": 0.001,
    "KILOGRAMS": 1.0, "KG": 1.0,
    "POUNDS": 0.45359237, "LB": 0.45359237,
    "OUNCES": 0.028349523125, "OZ": 0.028349523125,
}


def to_float(value) -> float | None:
    try:
        if value in (None, ""):
            return None
        return float(str(value).replace(",", ""))
    except ValueError:
        return None


def products_from_graphql_dump(dump: dict) -> list[Product]:
    by_id: dict[str, Product] = {}
    order = []
    for n in dump["products"]:
        p = Product(
            product_id=str(n.get("legacyResourceId") or n["id"].rsplit("/", 1)[-1]),
            handle=n["handle"],
            title=n.get("title") or "",
            status=(n.get("status") or "").lower(),
            product_type=n.get("productType") or "",
            category=((n.get("category") or {}).get("fullName")) or "",
            vendor=n.get("vendor") or "",
            tags=list(n.get("tags") or []),
            body_html=n.get("descriptionHtml") or "",
            image_url=((((n.get("featuredMedia") or {}).get("preview") or {}).get("image") or {}).get("url")) or "",
            collections=[c["title"] for c in ((n.get("collections") or {}).get("nodes") or [])],
            metafields={
                f"{m['namespace']}.{m['key']}": (m.get("type") or "", m.get("value") or "")
                for m in ((n.get("metafields") or {}).get("nodes") or [])
            },
            variants_count=((n.get("variantsCount") or {}).get("count")),
        )
        by_id[n["id"]] = p
        order.append(p)
    for v in dump.get("variants", []):
        p = by_id.get(v["product"]["id"])
        if p is None:
            continue
        w = ((v.get("inventoryItem") or {}).get("measurement") or {}).get("weight") or {}
        kg = None
        if to_float(w.get("value")):
            kg = to_float(w["value"]) * WEIGHT_TO_KG.get(str(w.get("unit", "")).upper(), 1.0)
        p.variants.append(
            Variant(
                sku=(v.get("sku") or "").strip(),
                price=to_float(v.get("price")),
                weight_kg=kg,
                title=v.get("title") or "",
                options={o["name"]: o["value"] for o in (v.get("selectedOptions") or [])},
            )
        )
    return order


# --------------------------------------------------------------------------- #
# Source 3: Shopify product CSV export
# --------------------------------------------------------------------------- #
METAFIELD_COL_RE = re.compile(r"\(product\.metafields\.([^.)]+)\.([^)]+)\)")


def products_from_csv(path: Path) -> list[Product]:
    with path.open(encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        return []
    cols = rows[0].keys()
    meta_cols = {c: f"{m.group(1)}.{m.group(2)}" for c in cols if (m := METAFIELD_COL_RE.search(c))}
    has_status = "Status" in cols
    weight_unit_col = "Variant Weight Unit"

    products: dict[str, Product] = {}
    images: dict[str, list] = {}
    for r in rows:
        handle = (r.get("Handle") or "").strip()
        if not handle:
            continue
        p = products.get(handle)
        if p is None:
            p = products[handle] = Product(handle=handle)
            images[handle] = []
        # Product-level fields live on the first row of each handle.
        if r.get("Title") and not p.title:
            p.title = r["Title"].strip()
            p.body_html = r.get("Body (HTML)") or ""
            p.vendor = r.get("Vendor") or ""
            p.product_type = r.get("Type") or ""
            p.category = r.get("Product Category") or ""
            p.tags = [t.strip() for t in (r.get("Tags") or "").split(",") if t.strip()]
            p.status = (r.get("Status") or "").strip().lower() if has_status else "unknown"
            for c, key in meta_cols.items():
                if r.get(c):
                    p.metafields[key] = ("csv", r[c])
        if r.get("Image Src"):
            pos = to_float(r.get("Image Position")) or 9999
            images[handle].append((pos, r["Image Src"].strip()))
        is_variant = any(r.get(c) for c in ("Variant Price", "Variant SKU", "Option1 Value"))
        if is_variant:
            grams = to_float(r.get("Variant Grams"))
            kg = grams / 1000.0 if grams else None
            if kg is None and to_float(r.get("Variant Weight")):  # some third-party exports
                kg = to_float(r["Variant Weight"]) * WEIGHT_TO_KG.get(
                    (r.get(weight_unit_col) or "kg").upper(), 1.0
                )
            opts = {}
            for i in (1, 2, 3):
                name, val = r.get(f"Option{i} Name"), r.get(f"Option{i} Value")
                if val:
                    opts[name or f"Option{i}"] = val
            p.variants.append(
                Variant(
                    sku=(r.get("Variant SKU") or "").strip(),
                    price=to_float(r.get("Variant Price")),
                    weight_kg=kg,
                    title=" / ".join(opts.values()),
                    options=opts,
                )
            )
    for handle, p in products.items():
        if images[handle]:
            p.image_url = sorted(images[handle])[0][1]
    result = list(products.values())
    if has_status:
        result = [p for p in result if p.status == "active"]
    else:
        print("WARNING: CSV has no 'Status' column - including every product in the file.")
    return result


# --------------------------------------------------------------------------- #
# Text helpers
# --------------------------------------------------------------------------- #
def html_to_text(raw: str) -> str:
    if not raw:
        return ""
    s = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", raw)
    s = re.sub(r"(?i)</t[dh]>\s*<t[dh][^>]*>", " : ", s)  # table cell -> "Label : value"
    s = re.sub(r"(?i)<br\s*/?>|</(p|div|li|tr|h[1-6]|ul|ol|table)>", "\n", s)
    s = re.sub(r"<[^>]+>", " ", s)
    s = html.unescape(s)
    s = (
        s.replace("″", '"').replace("“", '"').replace("”", '"')
        .replace("′", "'").replace("’", "'").replace(" ", " ")
        .replace("×", "x").replace("✕", "x").replace("✖", "x")
    )
    s = re.sub(r"(\d),(\d{3})(?!\d)", r"\1\2", s)  # 1,200 -> 1200
    s = re.sub(r"(\d),(\d{1,2})(?!\d)", r"\1.\2", s)  # 51,5 -> 51.5
    s = re.sub(r"[ \t]+", " ", s)
    return s


def fmt_num(x: float) -> str:
    x = round(x, 1)
    return str(int(x)) if x == int(x) else f"{x:.1f}"


# --------------------------------------------------------------------------- #
# Material detection
# --------------------------------------------------------------------------- #
# (regex, canonical name, is_wood)
MATERIALS = [
    (r"\bmango\s*wood\b", "Mango wood", True),
    (r"\b(?:solid\s+|hard|soft|drift|rose|ply|bass|red|iron|teak)?wood(?:en|s)?\b|\btimber\b", "Wood", True),
    (r"\boak\b", "Oak", True),
    (r"\bpine\b", "Pine", True),
    (r"\bteak\b", "Teak", True),
    (r"\bwalnut\b", "Walnut", True),
    (r"\bacacia\b", "Acacia", True),
    (r"\bbeech\b", "Beech", True),
    (r"\bbirch\b", "Birch", True),
    (r"\bpaulownia\b", "Paulownia", True),
    (r"\bbamboo\b", "Bamboo", True),
    (r"\bmdf\b", "MDF", True),
    (r"\bply\s?wood\b|\bplywood\b", "Plywood", True),
    (r"\b(?:chip|particle)\s?board\b", "Particleboard", True),
    (r"\brattan\b", "Rattan", True),
    (r"\bwicker\b", "Wicker", True),
    (r"(?<!candy )\bcane\b", "Cane", True),
    (r"\bmetal\b", "Metal", False),
    (r"\b(?:wrought |cast )?iron\b", "Iron", False),
    (r"\b(?:stainless )?steel\b", "Steel", False),
    (r"\balumini?um\b", "Aluminium", False),
    (r"\bbrass\b", "Brass", False),
    (r"\bcopper\b", "Copper", False),
    (r"\bzinc\b", "Zinc", False),
    (r"\bglass\b", "Glass", False),
    (r"\bcrystal\b", "Crystal", False),
    (r"\bceramic\b", "Ceramic", False),
    (r"\bporcelain\b", "Porcelain", False),
    (r"\bstoneware\b", "Stoneware", False),
    (r"\bterracotta\b|\bterra cotta\b", "Terracotta", False),
    (r"\bmarble\b", "Marble", False),
    (r"\bconcrete\b|\bcement\b", "Concrete", False),
    (r"\bfabric\b", "Fabric", False),
    (r"\bcotton\b", "Cotton", False),
    (r"\blinen\b", "Linen", False),
    (r"\bvelvet\b", "Velvet", False),
    (r"\bpolyester\b", "Polyester", False),
    (r"\bwool\b", "Wool", False),
    (r"\bjute\b", "Jute", False),
    (r"\bseagrass\b", "Seagrass", False),
    (r"\b(?:faux |pu |vegan )?leather\b", "Leather/PU", False),
    (r"\bplastic\b|\bpolypropylene\b|\bpp\b|\bpvc\b|\babs\b", "Plastic", False),
    (r"\b(?:poly)?resin\b", "Resin", False),
    (r"\beva\b", "EVA foam", False),
    (r"\b(?:memory )?foam\b", "Foam", False),
    (r"\bmicrofib(?:re|er)\b", "Microfibre", False),
    (r"\bnylon\b", "Nylon", False),
    (r"\brubber\b", "Rubber", False),
    (r"\bpolycarbonate\b|\bpc plastic\b", "Polycarbonate", False),
    (r"\bacrylic\b", "Acrylic", False),
    (r"\bsilicone\b", "Silicone", False),
]
MATERIALS = [(re.compile(rx, re.I), name, wood) for rx, name, wood in MATERIALS]
WOOD_NAMES = {name for _, name, wood in MATERIALS if wood}

# "Material: Aluminum + Iron" / "Material : Oak" (spec tables become "Label : value")
MATERIAL_FIELD_RE = re.compile(
    r"(?im)^\s*(?:main\s+|frame\s+|shade\s+|body\s+|top\s+|base\s+|product\s+)?materials?(?:\s*\([^)]*\))?\s*[:\-]\s*(.+)$"
)


def find_materials(text: str) -> list[str]:
    found = []
    for rx, name, _ in MATERIALS:
        if rx.search(text) and name not in found:
            found.append(name)
    if "EVA foam" in found and "Foam" in found:
        found.remove("Foam")
    # "Wood" is redundant when a specific wood was named
    if "Wood" in found and len([f for f in found if f in WOOD_NAMES]) > 1:
        found.remove("Wood")
    return found


def detect_materials(p: Product, text: str) -> tuple[list[str], str]:
    """Return (materials, source). Structured sources win over free text."""
    meta_vals = [
        v for k, (_, v) in p.metafields.items()
        if re.search(r"material|fabric|finish|composition", k, re.I)
    ]
    if meta_vals:
        mats = find_materials(" ; ".join(meta_vals))
        if mats:
            return mats, "metafield"
    field_vals = [m.group(1) for m in MATERIAL_FIELD_RE.finditer(text)]
    if field_vals:
        mats = find_materials(" ; ".join(field_vals))
        if mats:
            return mats, "description 'Material' field"
    mats = find_materials(" ; ".join(p.tags))
    if mats:
        return mats, "tags"
    mats = find_materials(strip_surface_mentions(p.title + " \n " + text))
    if mats:
        return mats, "title/description keywords"
    return [], ""


# "grips tiles, timber and other hard floors" - describes where it is used, not what it is made of
SURFACE_CONTEXT_RE = re.compile(r"(?i)\bfloor|\btiles?\b|\blaminate\b|\bbrick\b|\bdeck(?:ing)?\b")
# "wood-look", "wood-grain finish", "Colour options: ... Wood finish" - appearance only
LOOK_AFTER_RE = re.compile(r"(?i)^[\s-]*(?:look|grain|effect|finish|tones?|colou?r|print|pattern|style)\b")
COLOUR_LIST_RE = re.compile(r"(?i)colou?r(?:\s+options?)?\s*:[^\n]*$")


def strip_surface_mentions(text: str) -> str:
    """Remove wood words that describe the surface a product is used on ("tiles and timber floors")."""
    for rx, _, wood in MATERIALS:
        if not wood:
            continue
        def _drop(m, _t=text):
            line_start = _t.rfind("\n", 0, m.start()) + 1
            before = _t[max(line_start, m.start() - 15): m.start()]
            after = _t[m.end(): m.end() + 30]
            return " " if SURFACE_CONTEXT_RE.search(before + " " + after) else m.group(0)
        text = rx.sub(_drop, text)
    return text


def wood_evidence(p: Product, text: str, materials: list[str]) -> tuple[list[str], list[str]]:
    """Return (strong, weak) wood mentions.

    strong = the product appears to contain wood; weak = only a wood look/finish/colour is mentioned.
    Mentions of the surface the product sits on (timber floors, etc.) are ignored.
    """
    metafield_text = " ; ".join(v for k, (_, v) in p.metafields.items() if not k.startswith("global."))
    hay = "\n".join([p.title, text, " ; ".join(p.tags), metafield_text])
    strong, weak = [], []
    for rx, name, wood in MATERIALS:
        if not wood:
            continue
        for m in rx.finditer(hay):
            word = m.group(0).lower()
            line_start = hay.rfind("\n", 0, m.start()) + 1
            before = hay[max(line_start, m.start() - 30): m.start()]
            after = hay[m.end(): m.end() + 30]
            if SURFACE_CONTEXT_RE.search(before[-15:] + " " + after):
                continue
            if LOOK_AFTER_RE.match(after) or COLOUR_LIST_RE.search(hay[line_start: m.start()]):
                if word not in weak:
                    weak.append(word)
                continue
            if word not in strong:
                strong.append(word)
    if not strong:
        for mat in materials:
            if mat in WOOD_NAMES and mat.lower() not in strong:
                strong.append(mat.lower())
    weak = [w for w in weak if w not in strong]
    return strong, weak


# --------------------------------------------------------------------------- #
# Dimension parsing
# --------------------------------------------------------------------------- #
NUM = r"\d+(?:\.\d+)?"
QUAL = r"(?:(?:approx(?:imately)?\.?|about|around|ca\.|~|up to|max\.?)[ \t]*)?"
UNIT = (
    r"(?:mm|millimet(?:er|re)s?|cm|centimet(?:er|re)s?|met(?:er|re)s?|m(?![a-z])"
    r"|inch(?:es)?|in(?![a-z])|''|\"|ft|feet|foot)"
)
UNIT_TO_CM = {"mm": 0.1, "cm": 1.0, "m": 100.0, "in": 2.54, "ft": 30.48}
METRIC = {"mm", "cm", "m"}

PREFIX_LABEL = r"(?:Length|Width|Depth|Height|Diameter|Dia\.?|Thickness|Ø|L|W|D|H)"
SUFFIX_LABEL = r"(?:\((?:L|W|D|H|Ø)\)|L|W|D|H)(?![a-z])"
COMPONENT = (
    rf"(?:(?P<pre{{i}}>{PREFIX_LABEL})\s*[:.]?\s*)?(?P<num{{i}}>{NUM})\s*(?P<unit{{i}}>{UNIT})?"
    rf"\s*(?P<suf{{i}}>{SUFFIX_LABEL})?"
)
SEP = r"[ \t]*[xX*][ \t]*"
GROUP_RE = re.compile(
    r"(?<![\w.])"
    + COMPONENT.format(i=1) + SEP + COMPONENT.format(i=2)
    + rf"(?:{SEP}" + COMPONENT.format(i=3) + r")?",
    re.I,
)
SINGLE_RE = re.compile(
    rf"(?P<lab>\b(?:Length|Width|Depth|Height|Diameter|Dia|Thickness)\b|Ø|\b[LWDH](?=\s*[:=]?\s*\d))"
    rf"[ \t]*(?:\((?:cm|mm|in|inches)\))?[ \t]*[:=\-]?[ \t]*{QUAL}(?P<num>{NUM})[ \t]*(?P<unit>{UNIT})",
    re.I,
)
SINGLE_SUFFIX_RE = re.compile(
    rf"(?<![\w.])(?P<num>{NUM})[ \t]*(?P<unit>{UNIT})[ \t]*(?P<lab>high|tall|wide|deep|long|across|in diameter|diameter)\b(?![ \t]*:)",
    re.I,
)
SIZE_FIELD_RE = re.compile(
    rf"(?i)\b(?:dimensions?|size|measurements?|overall)\b[^:\n]{{0,25}}:\s*{QUAL}(?P<num>{NUM})\s*(?P<unit>{UNIT})(?!\s*[xX*]\s*\d)"
)
ORDER_HINT_RE = re.compile(
    r"[ \t]*\((\s*(?:length|width|depth|height|diameter|[LWDH])(?:\s*[xX*]\s*(?:length|width|depth|height|diameter|[LWDH])){1,2}\s*)\)",
    re.I,
)
UNIT_CONTEXT_RE = re.compile(r"(?i)\((cm|mm|inch(?:es)?|in)\)|\bin (cm|mm|inches)\b|\b(cm|mm)\s*:")
DIM_LABEL_CONTEXT_RE = re.compile(r"(?i)dimension|size|measure")
PACKAGE_CONTEXT_RE = re.compile(r"(?i)packag|carton|box size|shipping|parcel|outer box|gross")


def unit_key(u: str | None) -> str | None:
    if not u:
        return None
    u = u.lower().strip()
    if u in ('"', "''") or u.startswith("in"):
        return "in"
    if u in ("ft", "feet", "foot"):
        return "ft"
    if u.startswith("mm") or u.startswith("milli"):
        return "mm"
    if u.startswith("cm") or u.startswith("centi"):
        return "cm"
    if u.startswith("m"):
        return "m"
    return None


def axis_of(label: str | None) -> str | None:
    if not label:
        return None
    l = label.strip("().: ").lower()
    if l in ("ø", "dia", "dia.", "diameter", "in diameter", "across"):
        return "DIA"
    return {
        "l": "L", "length": "L", "long": "L",
        "w": "W", "width": "W", "wide": "W",
        "d": "D", "depth": "D", "deep": "D",
        "h": "H", "height": "H", "high": "H", "tall": "H",
        "thickness": "T",
    }.get(l)


@dataclass
class DimSet:
    values: dict  # axis -> cm ; axes: L W D H DIA T and U1.. (unlabelled, in order)
    metric: bool
    assumed_unit: bool = False
    package: bool = False
    raw: str = ""

    def lwh(self) -> tuple:
        v = self.values
        unl = [v[k] for k in sorted(k for k in v if k.startswith("U"))]
        if unl and not any(k in v for k in ("L", "W", "D", "H", "DIA", "T")):
            # unlabelled "a x b x c": keep the order as written (L x W x H)
            vals = unl + [None] * (3 - len(unl))
            return tuple(vals[:3]), False
        horiz = [v[k] for k in ("L", "W", "D") if k in v]
        dia = "DIA" in v
        if dia:
            horiz += [v["DIA"], v["DIA"]]
        horiz += unl
        horiz.sort(reverse=True)
        h = v.get("H", v.get("T"))
        L = horiz[0] if horiz else None
        W = horiz[1] if len(horiz) > 1 else None
        return (L, W, h), dia

    def max_dim(self) -> float:
        return max(self.values.values())

    def volume_key(self) -> float:
        prod = 1.0
        for x in self.values.values():
            prod *= x
        return prod

    def key(self) -> tuple:
        return tuple(sorted(round(x) for x in self.values.values()))


def _context_unit(text: str, start: int) -> tuple[str | None, bool]:
    before = text[max(0, start - 80): start]
    m = None
    for m in UNIT_CONTEXT_RE.finditer(before):
        pass
    if m:
        return unit_key(next(g for g in m.groups() if g)), False
    if DIM_LABEL_CONTEXT_RE.search(text[max(0, start - 40): start]):
        return "cm", True
    return None, False


def parse_dimensions(text: str) -> list[DimSet]:
    sets: list[DimSet] = []
    if not text:
        return sets
    covered = []
    for m in GROUP_RE.finditer(text):
        comps = []
        for i in (1, 2, 3):
            if m.group(f"num{i}") is None:
                continue
            comps.append(
                {
                    "num": float(m.group(f"num{i}")),
                    "unit": unit_key(m.group(f"unit{i}")),
                    "axis": axis_of(m.group(f"pre{i}")) or axis_of(m.group(f"suf{i}")),
                }
            )
        units = [c["unit"] for c in comps if c["unit"]]
        assumed = False
        if units:
            default = units[-1]
        else:
            default, assumed = _context_unit(text, m.start())
            if default is None:
                continue  # "2 x 3" with no unit anywhere -> not a dimension
        # an order hint right after the numbers: "86.5 x 23.8 cm (height x width)", "(W x D x H)"
        hint = ORDER_HINT_RE.match(text, m.end())
        if hint and all(c["axis"] is None for c in comps):
            axes = [axis_of(a.strip()) for a in re.split(r"\s*[xX*]\s*", hint.group(1))]
            if len(axes) == len(comps) and all(axes) and len(set(axes)) == len(axes):
                for c, a in zip(comps, axes):
                    c["axis"] = a
        values = {}
        n_unl = 0
        for c in comps:
            u = c["unit"] or default
            cm = c["num"] * UNIT_TO_CM[u]
            if cm <= 0:
                continue
            axis = c["axis"]
            if axis is None or axis in values:
                n_unl += 1
                axis = f"U{n_unl}"
            values[axis] = cm
        if len(values) < 2:
            continue
        line_start = text.rfind("\n", 0, m.start()) + 1
        ds = DimSet(
            values=values,
            metric=default in METRIC and all(c["unit"] in (None, *METRIC) for c in comps),
            assumed_unit=assumed,
            package=bool(PACKAGE_CONTEXT_RE.search(text[max(line_start, m.start() - 60): m.start()])),
            raw=m.group(0).strip(),
        )
        sets.append(ds)
        covered.append((m.start(), m.end()))

    def is_covered(pos):
        return any(a <= pos < b for a, b in covered)

    # Labelled single values ("Height: 45cm", "Ø 30 cm", "45cm high") grouped per line.
    singles_by_line: dict[int, dict] = {}
    for rx in (SINGLE_RE, SINGLE_SUFFIX_RE):
        for m in rx.finditer(text):
            if is_covered(m.start()):
                continue
            u = unit_key(m.group("unit"))
            axis = axis_of(m.group("lab"))
            if not u or not axis:
                continue
            line_start = text.rfind("\n", 0, m.start()) + 1
            # all singles from one spec block count as one set: key on the paragraph
            block = text.rfind("\n\n", 0, m.start())
            entry = singles_by_line.setdefault(block, {"values": {}, "metric": set(), "pkg": False, "raw": []})
            cm = float(m.group("num")) * UNIT_TO_CM[u]
            if axis in entry["values"]:
                entry["values"][axis] = max(entry["values"][axis], cm)
            else:
                entry["values"][axis] = cm
            entry["metric"].add(u in METRIC)
            entry["raw"].append(m.group(0).strip())
            if PACKAGE_CONTEXT_RE.search(text[max(line_start, m.start() - 60): m.start()]):
                entry["pkg"] = True
    for entry in singles_by_line.values():
        sets.append(
            DimSet(
                values=entry["values"],
                metric=entry["metric"] == {True},
                package=entry["pkg"],
                raw="; ".join(entry["raw"]),
            )
        )
    # "Dimensions : 51.5 cm" - one overall number, axis not stated
    if not sets:
        for m in SIZE_FIELD_RE.finditer(text):
            u = unit_key(m.group("unit"))
            if u:
                sets.append(
                    DimSet(
                        values={"U1": float(m.group("num")) * UNIT_TO_CM[u]},
                        metric=u in METRIC,
                        raw=m.group(0).strip(),
                    )
                )
    # If the text gives metric figures, drop imperial duplicates such as (Ø5" x H7")
    if any(s.metric for s in sets):
        sets = [s for s in sets if s.metric]
    return sets


def parse_dimension_metafields(p: Product) -> list[DimSet]:
    sets = []
    single = {}
    for k, (mtype, val) in p.metafields.items():
        key = k.split(".", 1)[-1].lower()
        if not re.search(r"dimension|size|height|width|length|depth|diameter|measure", key):
            continue
        if not val:
            continue
        # Shopify "dimension" metafield: {"value": 45, "unit": "cm"} (list types are JSON arrays)
        parsed = None
        if isinstance(val, str) and val.strip()[:1] in "[{":
            try:
                parsed = json.loads(val)
            except ValueError:
                parsed = None
        if parsed is not None:
            items = parsed if isinstance(parsed, list) else [parsed]
            for it in items:
                if isinstance(it, dict) and "value" in it:
                    u = unit_key(str(it.get("unit", "cm")).lower().replace("centimeters", "cm")
                                 .replace("millimeters", "mm").replace("inches", "in").replace("meters", "m"))
                    axis = axis_of(re.sub(r".*?(height|width|length|depth|diameter).*", r"\1", key))
                    if u and axis and to_float(it["value"]):
                        single[axis] = max(single.get(axis, 0), to_float(it["value"]) * UNIT_TO_CM[u])
            continue
        found = parse_dimensions(html_to_text(str(val)))
        if found:
            sets.extend(found)
            continue
        # bare number in a key like custom.height - assume cm
        axis = axis_of(re.sub(r".*?(height|width|length|depth|diameter).*", r"\1", key))
        num = to_float(val)
        if axis and num:
            single[axis] = max(single.get(axis, 0), num)
    if single:
        sets.append(DimSet(values=single, metric=True, raw="metafields"))
    return sets


SIZE_OPTION_RE = re.compile(r"(?i)size|dimension")


def detect_dimensions(p: Product, text: str) -> tuple[str, list | None, str, bool]:
    """Return (display string, [L, W, H] in cm or None, source, found_any)."""
    source = "metafield"
    sets = parse_dimension_metafields(p)
    if not sets:
        source = "description"
        sets = parse_dimensions(text)
    variant_sets = []
    for v in p.variants:
        variant_sets.extend(parse_dimensions(html_to_text(v.title)))
    if not sets and variant_sets:
        source = "variant options"
    if variant_sets:
        if any(s.metric for s in variant_sets):
            variant_sets = [s for s in variant_sets if s.metric]
        sets = sets + variant_sets
    if not sets:
        sets = parse_dimensions(html_to_text(p.title))
        source = "title"
    if not sets:
        return "", None, "", False

    product_sets = [s for s in sets if not s.package] or sets
    only_package = all(s.package for s in product_sets)
    best = max(product_sets, key=lambda s: (s.max_dim(), s.volume_key()))
    (L, W, H), dia = best.lwh()
    parts = [fmt_num(x) if x is not None else "?" for x in (L, W, H)]
    display = " x ".join(parts)
    notes = []
    if dia:
        notes.append(f"Ø{fmt_num(best.values['DIA'])} round")
    if list(best.values) == ["U1"]:
        display = fmt_num(best.values["U1"])
        notes.append("single measurement, axis not stated")

    distinct = {s.key() for s in product_sets}
    size_values = {
        val for v in p.variants for name, val in v.options.items() if SIZE_OPTION_RE.search(name or "")
    }
    if len(distinct) > 1 or len(size_values) > 1:
        notes.append("multiple sizes - largest shown")
    if best.assumed_unit:
        notes.append("unit not stated, assumed cm")
    if only_package:
        notes.append("packaging size")
    if notes:
        display += " (" + "; ".join(notes) + ")"
    lwh = [x for x in (L, W, H)]
    return display, lwh, source, True


WEIGHT_TEXT_RE = re.compile(
    rf"(?i)\b(?:net\s+weight|gross\s+weight|item\s+weight|product\s+weight|weight|n\.?w\.?|g\.?w\.?)"
    rf"(?:\s*\([^)]*\))?\s*[:\-=]?\s*(?P<num>{NUM})\s*(?P<unit>kgs?|kilograms?|g|grams?|lbs?|pounds?|oz|ounces?)\b"
)


def detect_weight(p: Product, text: str) -> tuple[float | None, str]:
    weights = [v.weight_kg for v in p.variants if v.weight_kg]
    if weights:
        return max(weights), "variant weight"
    found = []
    for m in WEIGHT_TEXT_RE.finditer(text):
        u = m.group("unit").lower()
        factor = 1.0 if u.startswith("k") else 0.001 if u.startswith("g") else (
            0.45359237 if u.startswith(("lb", "pound")) else 0.028349523125)
        found.append(float(m.group("num")) * factor)
    if found:
        return max(found), "description"
    return None, ""


# --------------------------------------------------------------------------- #
# Category inference
# --------------------------------------------------------------------------- #
GENERIC_COLLECTIONS = re.compile(
    r"(?i)^(home ?page|frontpage|all|all products|shop all|sale|new|new arrivals|best ?sellers?|"
    r"featured|trending|gift.*|under \$?\d+.*|clearance|deals?)$"
)
CATEGORY_KEYWORDS = [
    (r"pendant|chandelier|ceiling light|wall light|sconce|lamp|lantern|light\b|lighting", "Lighting"),
    (r"\bmirror", "Mirrors"),
    (r"\brug\b|\brugs\b|\bmat\b|doormat", "Rugs & mats"),
    (r"cushion|pillow", "Cushions"),
    (r"\bthrow\b|blanket", "Throws & blankets"),
    (r"quilt|duvet|sheet set|bedding|bedspread|comforter", "Bedding"),
    (r"curtain|drape|blind", "Window furnishings"),
    (r"\bvase", "Vases"),
    (r"planter|plant pot|\bpots?\b", "Planters & pots"),
    (r"artificial|faux (?:plant|flower|tree)|\bstems?\b", "Artificial plants & flowers"),
    (r"candle|diffuser|fragrance", "Candles & fragrance"),
    (r"\bclock", "Clocks"),
    (r"wall art|canvas|print\b|poster|frame", "Wall art & frames"),
    (r"basket|storage|organi[sz]er|\bbox\b|hamper", "Storage & baskets"),
    (r"sofa|couch|armchair|chair|stool|bench|ottoman|pouf", "Seating"),
    (r"table|desk|sideboard|console|cabinet|dresser|drawers|bookcase|shel(?:f|ves)|tv unit|bed frame", "Furniture"),
    (r"\btray|bowl|platter|plate|mug|cup|glassware|serving|cutlery|kitchen|chopping|cutting board", "Kitchen & dining"),
    (r"towel|bath mat|soap|bathroom", "Bathroom"),
    (r"ornament|sculpture|figurine|statue|bookend|decor", "Decor & ornaments"),
    (r"outdoor|garden|patio", "Outdoor & garden"),
    (r"christmas|xmas|easter|halloween", "Seasonal"),
]
CATEGORY_KEYWORDS = [(re.compile(rx, re.I), name) for rx, name in CATEGORY_KEYWORDS]


def detect_category(p: Product) -> str:
    if p.product_type.strip():
        return p.product_type.strip()
    if p.category.strip():
        return p.category.split(">")[-1].strip() + " (inferred)"
    for c in p.collections:
        if not GENERIC_COLLECTIONS.match(c.strip()):
            return c.strip() + " (inferred)"
    hay = p.title + " " + " ".join(p.tags)
    for rx, name in CATEGORY_KEYWORDS:
        if rx.search(hay):
            return name + " (inferred)"
    return "Unclear"


# --------------------------------------------------------------------------- #
# Row building
# --------------------------------------------------------------------------- #
def build_rows(products: list[Product], domain: str) -> list[dict]:
    rows = []
    for p in products:
        text = html_to_text(p.body_html)
        materials, mat_src = detect_materials(p, text)
        wood_hits, wood_weak = wood_evidence(p, text, materials)
        if wood_hits:
            contains_wood = "Yes"
        elif wood_weak:
            contains_wood = "Unclear"  # only a wood look/finish/colour is mentioned
        elif materials:
            contains_wood = "No"
        else:
            contains_wood = "Unclear"

        dims_display, lwh, dims_src, dims_found = detect_dimensions(p, text)
        weight, weight_src = detect_weight(p, text)

        max_dim = max([x for x in (lwh or []) if x is not None], default=None)
        if (max_dim is not None and max_dim > OVERSIZE_CM) or (weight is not None and weight > OVERSIZE_KG):
            oversized = "Yes"
        elif not dims_found:
            oversized = "Unclear"
        else:
            oversized = "No"

        prices = [v.price for v in p.variants if v.price is not None]
        if prices:
            lo, hi = min(prices), max(prices)
            price = f"{lo:.2f}" if lo == hi else f"{lo:.2f} - {hi:.2f}"
        else:
            price = ""
        skus = []
        for v in p.variants:
            if v.sku and v.sku not in skus:
                skus.append(v.sku)

        n_variants = p.variants_count if p.variants_count is not None else len(p.variants)
        rows.append(
            {
                COLUMNS[0]: f"https://{domain}/products/{p.handle}",
                COLUMNS[1]: p.title,
                COLUMNS[2]: detect_category(p),
                COLUMNS[3]: ", ".join(materials) if materials else "Unclear",
                COLUMNS[4]: dims_display or "Unclear",
                COLUMNS[5]: round(weight, 3) if weight is not None else None,
                COLUMNS[6]: n_variants,
                COLUMNS[7]: ", ".join(skus),
                COLUMNS[8]: price,
                COLUMNS[9]: p.image_url,
                COLUMNS[10]: contains_wood,
                COLUMNS[11]: oversized,
                COLUMNS[12]: "",
                COLUMNS[13]: "",
                COLUMNS[14]: "",
                "_internal": {
                    "product_id": p.product_id,
                    "handle": p.handle,
                    "status": p.status,
                    "collections": ", ".join(p.collections),
                    "title": p.title,
                    "vendor": p.vendor,
                    "wood_evidence": ", ".join(wood_hits)
                    + (("; look/finish only: " if wood_hits else "look/finish only: ") + ", ".join(wood_weak) if wood_weak else ""),
                    "material_source": mat_src,
                    "dimensions_source": dims_src,
                    "weight_source": weight_src,
                },
                "_dims_unclear": not dims_found,
                "_material_unclear": not materials,
            }
        )

    def group(r):
        if r[COLUMNS[10]] == "Yes" or r[COLUMNS[11]] == "Yes":
            return 0
        if r[COLUMNS[10]] == "Unclear" or r[COLUMNS[11]] == "Unclear":
            return 1
        return 2

    rows.sort(key=lambda r: (group(r), r[COLUMNS[1]].lower()))
    for r in rows:
        r["_group"] = group(r)
    return rows


# --------------------------------------------------------------------------- #
# Output
# --------------------------------------------------------------------------- #
def write_csv(rows: list[dict], path: Path) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(COLUMNS)
        for r in rows:
            w.writerow(["" if r[c] is None else r[c] for c in COLUMNS])


def write_xlsx(rows: list[dict], path: Path) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation

    red = PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE")
    yellow = PatternFill("solid", start_color="FFEB9C", end_color="FFEB9C")
    green = PatternFill("solid", start_color="C6EFCE", end_color="C6EFCE")
    grey = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
    bold = Font(bold=True)
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    top_wrap = Alignment(vertical="top", wrap_text=True)
    top = Alignment(vertical="top")

    wb = Workbook()
    ws = wb.active
    ws.title = "Product review"
    ws.append(COLUMNS)
    for idx, name in enumerate(COLUMNS, start=1):
        c = ws.cell(row=1, column=idx)
        c.font = bold
        c.fill = green if name in WINWIN_COLUMNS else grey
        c.alignment = Alignment(vertical="center", wrap_text=True)
        c.border = border
    ws.row_dimensions[1].height = 45

    widths = [48, 42, 24, 26, 34, 11, 10, 32, 18, 48, 12, 12, 20, 32, 32]
    wrap_cols = {2, 3, 4, 5, 8, 14, 15}  # 1-based
    for i, wdt in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = wdt

    for r_idx, r in enumerate(rows, start=2):
        fill = red if r["_group"] == 0 else yellow if r["_group"] == 1 else None
        for c_idx, name in enumerate(COLUMNS, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=r[name])
            cell.alignment = top_wrap if c_idx in wrap_cols else top
            cell.border = border
            if fill is not None and name not in WINWIN_COLUMNS:
                cell.fill = fill
            if c_idx in (1, 10) and r[name]:
                cell.hyperlink = r[name]
                cell.font = Font(color="0563C1", underline="single")
            if c_idx == 6 and r[name] is not None:
                cell.number_format = "0.00"

    last_row = max(2, len(rows) + 1)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{last_row}"
    dv = DataValidation(type="list", formula1='"Yes,No,Unsure"', allow_blank=True)
    dv.error = "Please choose Yes, No or Unsure"
    dv.prompt = "Can this product ship to Australia via normal channels?"
    ws.add_data_validation(dv)
    dv.add(f"M2:M{last_row}")

    # Legend sheet-level note under the table is awkward with autofilter; use a 2nd sheet.
    info = wb.create_sheet("How to fill in")
    for line in [
        ["Hunter Valley Home - product review for Winwin"],
        [],
        ["Please fill in the three GREEN columns on the 'Product review' sheet:"],
        ["  Can ship normally to AU?", "Yes / No / Unsure (dropdown)"],
        ["  Reason if not", "e.g. wood - needs fumigation/treatment certificate, oversized, battery, restricted material"],
        ["  Notes", "anything else we should know"],
        [],
        ["Row colours"],
        ["  Light red", "Contains wood = Yes, or Oversized = Yes"],
        ["  Light yellow", "Wood or size could not be determined from our listing (Unclear)"],
        [],
        ["Oversized rule", f"any dimension > {OVERSIZE_CM:g} cm or weight > {OVERSIZE_KG:g} kg"],
        ["Dimensions", "Largest size shown when the product comes in several sizes; '?' = not stated"],
        ["Prices", "Current selling price on huntervalleyhome.com in AUD (min - max across variants)"],
    ]:
        info.append(line)
    info["A1"].font = Font(bold=True, size=13)
    info["A3"].font = bold
    info["A8"].font = bold
    info["A9"].fill = red
    info["A10"].fill = yellow
    info.column_dimensions["A"].width = 30
    info.column_dimensions["B"].width = 95

    internal = wb.create_sheet("Internal")
    internal.append(INTERNAL_COLUMNS)
    for idx in range(1, len(INTERNAL_COLUMNS) + 1):
        internal.cell(row=1, column=idx).font = bold
        internal.cell(row=1, column=idx).fill = grey
    for r in rows:
        internal.append([r["_internal"][c] for c in INTERNAL_COLUMNS])
    for i, wdt in enumerate([16, 45, 10, 50, 42, 18, 24, 26, 18, 16], start=1):
        internal.column_dimensions[get_column_letter(i)].width = wdt
    internal.freeze_panes = "A2"
    internal.auto_filter.ref = f"A1:{get_column_letter(len(INTERNAL_COLUMNS))}{max(2, len(rows) + 1)}"

    wb.save(path)


def print_summary(rows: list[dict], source: str) -> None:
    total = len(rows)
    wood = sum(r[COLUMNS[10]] == "Yes" for r in rows)
    oversized = sum(r[COLUMNS[11]] == "Yes" for r in rows)
    dims_unclear = sum(r["_dims_unclear"] for r in rows)
    mat_unclear = sum(r["_material_unclear"] for r in rows)
    flagged = sum(r["_group"] == 0 for r in rows)
    unclear_rows = sum(r["_group"] == 1 for r in rows)
    print()
    print("=" * 48)
    print(f" Hunter Valley Home product review  (source: {source})")
    print("=" * 48)
    print(f" Total active products      : {total}")
    print(f" Contains wood (Yes)        : {wood}")
    print(f" Oversized (Yes)            : {oversized}")
    print(f" Unclear dimensions         : {dims_unclear}")
    print(f" Unclear material           : {mat_unclear}")
    print("-" * 48)
    print(f" Red rows (wood/oversized)  : {flagged}")
    print(f" Yellow rows (unclear)      : {unclear_rows}")
    print(f" Clear rows                 : {total - flagged - unclear_rows}")
    print("=" * 48)


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--source", choices=["auto", "api", "json", "csv"], default="auto")
    ap.add_argument("--input-csv", type=Path, default=DEFAULT_INPUT_CSV)
    ap.add_argument("--input-json", type=Path, default=Path("input/products_graphql.json"))
    ap.add_argument("--save-json", type=Path, help="also save the raw API pull here (for re-runs without the API)")
    ap.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    ap.add_argument("--domain", default=STORE_DOMAIN)
    args = ap.parse_args(argv)

    load_dotenv([Path.cwd() / ".env", Path(__file__).resolve().parent / ".env"])
    store = os.environ.get("SHOPIFY_STORE", "").strip()
    token = os.environ.get("SHOPIFY_ADMIN_TOKEN", "").strip()
    api_version = os.environ.get("SHOPIFY_API_VERSION", DEFAULT_API_VERSION).strip()

    source = args.source
    if source == "auto":
        if store and token:
            source = "api"
        elif args.input_json.is_file():
            source = "json"
        elif args.input_csv.is_file():
            source = "csv"
        else:
            print(
                "No data source found. Provide one of:\n"
                "  - SHOPIFY_STORE and SHOPIFY_ADMIN_TOKEN in .env (Admin API, read_products scope)\n"
                f"  - a Shopify product export at {args.input_csv}\n"
                f"  - a GraphQL JSON dump at {args.input_json}",
                file=sys.stderr,
            )
            return 2

    if source == "api":
        if not (store and token):
            print("SHOPIFY_STORE / SHOPIFY_ADMIN_TOKEN missing from .env", file=sys.stderr)
            return 2
        dump = fetch_from_api(store, token, api_version)
        if dump["shop"].get("currencyCode") != "AUD":
            print(f"WARNING: store currency is {dump['shop'].get('currencyCode')}, not AUD.")
        if args.save_json:
            args.save_json.parent.mkdir(parents=True, exist_ok=True)
            args.save_json.write_text(json.dumps(dump), encoding="utf-8")
        products = products_from_graphql_dump(dump)
    elif source == "json":
        dump = json.loads(args.input_json.read_text(encoding="utf-8"))
        products = [p for p in products_from_graphql_dump(dump) if p.status in ("active", "")]
    else:
        if not args.input_csv.is_file():
            print(f"CSV not found: {args.input_csv}", file=sys.stderr)
            return 2
        products = products_from_csv(args.input_csv)
        print("Note: CSV exports carry no product IDs or collections - those Internal columns will be blank.")

    rows = build_rows(products, args.domain)
    args.output_dir.mkdir(parents=True, exist_ok=True)
    xlsx_path = args.output_dir / f"{OUTPUT_BASENAME}.xlsx"
    csv_path = args.output_dir / f"{OUTPUT_BASENAME}.csv"
    write_xlsx(rows, xlsx_path)
    write_csv(rows, csv_path)
    print_summary(rows, source)
    print(f" Wrote {xlsx_path}")
    print(f" Wrote {csv_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
