# Shopify Dev Dashboard Setup

Use this for the Can Count Sync app version.

## App URL

Use a placeholder for now if Shopify accepts it:

```text
https://example.com
```

Before going live with webhooks, replace it with your hosted app URL, for example:

```text
https://can-count-sync.onrender.com
```

Do not use your storefront URL.

## Embed app in Shopify admin

Leave off for V1.

## Preferences URL

Leave blank.

## Webhooks API version

Use:

```text
2026-04
```

## Scopes

Use exactly:

```text
write_inventory,read_inventory,read_locations,read_orders,read_products
```

## Optional scopes

Leave blank.

## Redirect URLs

Leave blank unless Shopify forces a value.

## POS

Leave off.

## App proxy

Leave blank.

## Webhook topics after hosting

Add these after the app is hosted:

```text
orders/create
products/create
products/update
```

Endpoint:

```text
https://YOUR-HOST/webhooks/shopify
```
