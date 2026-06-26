# Can Count Sync

Can Count Sync is a small Shopify inventory sync tool for Designated Drinks.

It does **not** rebuild the website. It does **not** manage prices. It does **not** change product pages.

It only does this:

```text
real cans available -> calculate Single / 4-Pack / 12-Pack availability -> sync Shopify inventory
```

## What Shopify owns

- product titles
- product images
- product prices
- pack option labels
- customer-facing product pages
- checkout

## What the sheet/app owns

- real can counts
- pack size mapping
- damaged/dropped/sample/received adjustments
- order deductions
- Shopify inventory sync numbers
- new variant registration queue

## Files

```text
can_count_sync_proper_template.xlsx   Google Sheet template
src/                                  app code
scripts/                              command-line scripts
docs/                                 setup instructions
.env.example                          environment variable example
```

## Commands

```bash
npm install
npm test
npm run locations
npm run import:variants
npm run promote:mapping
npm run sync:preview
npm run sync
npm run server
```

## Core math

If real cans = 237:

```text
Single inventory = floor(237 / 1) = 237
4-Pack inventory = floor(237 / 4) = 59
12-Pack inventory = floor(237 / 12) = 19
```

## New products

When new products or variants are added in Shopify:

1. Run `npm run import:variants`, or let the `products/create` / `products/update` webhook run it.
2. New variants go to `Needs Mapping`.
3. Confirm `proposed_base_sku` and `proposed_pack_size`.
4. Change status to `ready_to_promote`.
5. Run `npm run promote:mapping`.

No customer-facing rebuild required.
