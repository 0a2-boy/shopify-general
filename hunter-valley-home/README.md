# Hunter Valley Home

## Supplier product review sheet (Winwin, AU shipping)

`product_review.py` builds a one-row-per-product sheet of all **active** products so the
supplier Winwin can mark which ones can ship to Australia via normal channels.
It is **read-only**: it only sends GraphQL queries and refuses to send a mutation.

Run from the repo root:

```bash
python3 hunter-valley-home/product_review.py
```

Outputs:

- `output/hunter_valley_home_product_review.xlsx`, with three sheets: *Product review* (for Winwin), *How to fill in*, and *Internal* (product_id, handle, status, collections, plus where each value came from)
- `output/hunter_valley_home_product_review.csv`, the *Product review* sheet only

### Data source (picked automatically, or force with `--source api|json|csv`)

1. **Admin API**: put `SHOPIFY_STORE` and `SHOPIFY_ADMIN_TOKEN` in `.env` (see `.env.example`).
   The token needs only the `read_products` scope. Add `--save-json input/products_graphql.json`
   to cache the pull so you can re-run without the API.
2. **JSON dump**: `input/products_graphql.json`, in the same shape as the API pull.
3. **CSV export**: Shopify Admin → Products → Export → *All products*, *CSV for Excel*,
   saved as `input/products_export.csv`. The CSV has no product IDs or collections, so those
   Internal columns stay blank.

### How the columns are derived

| Column | Logic |
|---|---|
| Category | `product_type`. If that's empty: Shopify standard category, then the first non-generic collection, then title/tag keywords. Inferred values are marked "(inferred)". |
| Material(s) | First match wins: a material metafield, then the description's "Material:" spec row, then tags, then keywords in the title or description. |
| Dimensions | Dimension metafields, then the description, then variant option values. mm, m, inches and ft are converted to cm. Metric figures take precedence over inch duplicates. Packaging/carton sizes are ignored unless they are the only figures. When there are several sizes, the largest is shown with "multiple sizes". |
| Weight | Max variant weight (g/lb/oz converted to kg). 0 counts as missing. If no variant has a weight, falls back to a "Weight: x kg" line in the description. |
| Contains wood? | **Yes** if the materials, title, description, tags or metafields mention wood/wooden/timber/oak/pine/teak/walnut/acacia/beech/birch/mango wood/paulownia/bamboo/MDF/plywood/particleboard/rattan/wicker/cane. **No** if a material was found and none of these words appear. Otherwise **Unclear**. |
| Oversized? | **Yes** if any dimension > 60 cm or weight > 10 kg. **Unclear** if no dimensions were found. Otherwise **No**. |

Rows are sorted wood/oversized first (light red), then Unclear (light yellow), then the rest.
The three Winwin columns have green headers. "Can ship" has a Yes/No/Unsure dropdown.
