# Fennholt – payment icons

Order everywhere: Amex, Apple Pay, Google Pay, Maestro, Mastercard, Shop Pay, Visa.
Style: the Essential Cart Drawer icon set (70×48 cards, shown 24px high, 8px gap).
Mastercard is the Essential icon with a dark (#1F2730) background.

- **Icons** – `assets/payment-icons/fennholt-payment-*.svg`, uploaded to Shopify
  Settings > Files under the same names and loaded with `file_url`.
- **Footer** – "Custom liquid" block in `sections/footer-group.json`
  (`ai_gen_block_3515fa4`) renders the 7 images in the order above.
- **Cart page** – `snippets/cart-summary.liquid` renders `snippets/payment-icons-custom.liquid`
  below the checkout button.
- **Cart drawer** – rendered by the Essential Cart Drawer app, configured in the app
  (Payment icons).
