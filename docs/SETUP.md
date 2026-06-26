# Can Count Sync Setup

## 1. Create the Shopify app

Use Shopify Dev Dashboard.

App name:

```text
Can Count Sync
```

Scopes:

```text
write_inventory,read_inventory,read_locations,read_orders,read_products
```

Webhooks API version:

```text
2026-04
```

## 2. Create the Google Sheet

Upload the template spreadsheet to Google Drive and open it as a Google Sheet.

Rename it:

```text
Can Count Sync - Designated Drinks
```

Copy the Sheet ID from the browser URL and put it in `.env` as:

```text
GOOGLE_SHEET_ID=your_sheet_id
```

## 3. Add service account access

Create a Google service account JSON key.

Save it locally as:

```text
service-account.json
```

Share the Google Sheet with the service account email address as an editor.

## 4. Create `.env`

Copy `.env.example` to `.env` and fill in the real values.

Do not commit `.env`.

## 5. Install

```bash
npm install
npm test
```

## 6. Find Shopify location

```bash
npm run locations
```

Copy the correct location ID into `.env`.

## 7. Import variants

```bash
npm run import:variants
```

New Shopify variants appear in the `Needs Mapping` tab.

## 8. Map variants

For each variant, confirm:

```text
proposed_base_sku
proposed_pack_size
```

Set status to:

```text
ready_to_promote
```

Then run:

```bash
npm run promote:mapping
```

## 9. Add can counts

Use `Base Products` and `Can Ledger`.

Examples:

```text
+240 received
-3 damaged
-1 sample
```

## 10. Preview

```bash
npm run sync:preview
```

## 11. Live sync

```bash
npm run sync
```

This updates Shopify inventory only.
