# Fennholt – payment icons

Order everywhere: Amex, Apple Pay, Google Pay, Maestro, Mastercard, Shop Pay, Visa.

- **Footer** – "Custom liquid" block in `sections/footer-group.json`
  (`ai_gen_block_3515fa4`): `payment_types` list reordered, icons 38×24px, gap 10px.
- **Cart page** – `snippets/cart-summary.liquid` renders `snippets/payment-icons-custom.liquid`
  below the checkout button.
- **Cart drawer** – rendered by the Essential Cart Drawer app, configured in the app
  (Payment icons). The SVGs in `assets/payment-icons/` are the same Shopify icons as the
  footer and are also uploaded to Shopify Files as `fennholt-pay-*.svg`.
