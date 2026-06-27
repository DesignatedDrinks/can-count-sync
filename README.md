# Can Count Sync

Can Count Sync is a small Shopify inventory sync tool for Designated Drinks.

The recommended V1 is the **Apps Script Edition** inside Google Sheets.

It does **not** rebuild the website. It does **not** manage prices. It does **not** change product pages.

It only does this:

```text
real cans available -> calculate Single / 4-Pack / 12-Pack availability -> sync Shopify inventory
```

## Primary V1 path

Use:

```text
apps-script/Code.gs
```

Paste that file into:

```text
Google Sheet -> Extensions -> Apps Script
```

Then reload the sheet and use the **Can Count Sync** menu.

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
apps-script/  recommended Google Sheet version
src/          optional Node/server version for later
docs/         setup instructions
```

The Google Sheet template is provided separately as `can_count_sync_proper_template.xlsx`. Upload it to Google Drive and open it as a Google Sheet.

## Apps Script menu

```text
Can Count Sync
- Setup / Repair Sheet Tabs
- Import Shopify Variants
- Promote Ready Mapping Rows
- Preview Inventory Sync
- Sync Inventory to Shopify
- Process Recent Orders
- Install 10 Minute Order Polling
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

1. Run `Import Shopify Variants` from the Can Count Sync menu.
2. New variants go to `Needs Mapping`.
3. Confirm `proposed_base_sku` and `proposed_pack_size`.
4. Change status to `ready_to_promote`.
5. Run `Promote Ready Mapping Rows`.

No customer-facing rebuild required.
