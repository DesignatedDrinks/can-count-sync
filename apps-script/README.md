# Can Count Sync - Apps Script Edition

This is the simple V1.

It runs inside the Google Sheet. No Cloud Run. No Shopify CLI. No public server.

## Jobs

- import Shopify variants into Needs Mapping
- map each variant to a base can SKU
- track real cans in Base Products and Can Ledger
- calculate pack inventory
- sync Shopify inventory quantities
- poll recent Shopify orders and deduct cans

## It does not touch

- website theme
- product pages
- prices
- images
- product titles
- checkout design

## Install

Open the Google Sheet.

Go to Extensions, then Apps Script.

Create Code.gs.

Paste the Code.gs file from this folder.

Reload the sheet and use the Can Count Sync menu.
