import { google } from 'googleapis';
import fs from 'node:fs';
import { config } from './config.js';

export const SHEETS = {
  BASE_PRODUCTS: 'Base Products',
  SHOPIFY_VARIANTS: 'Shopify Variants',
  CAN_LEDGER: 'Can Ledger',
  NEEDS_MAPPING: 'Needs Mapping',
  SYNC_LOG: 'Sync Log',
  ORDERS_PROCESSED: 'Orders Processed'
};

export const HEADERS = {
  BASE_PRODUCTS: ['base_sku', 'product_name', 'case_size', 'opening_cans', 'safety_buffer_cans', 'active'],
  SHOPIFY_VARIANTS: ['base_sku', 'product_title', 'variant_title', 'sku', 'price', 'variant_gid', 'variant_legacy_id', 'inventory_item_id', 'pack_size', 'active', 'tracked', 'last_seen_at'],
  CAN_LEDGER: ['timestamp', 'base_sku', 'change_cans', 'reason', 'reference_type', 'reference_id', 'note', 'source'],
  NEEDS_MAPPING: ['status', 'product_title', 'variant_title', 'sku', 'price', 'proposed_base_sku', 'proposed_pack_size', 'variant_gid', 'variant_legacy_id', 'inventory_item_id', 'tracked', 'last_seen_at', 'note'],
  SYNC_LOG: ['timestamp', 'mode', 'base_sku', 'product_title', 'variant_title', 'pack_size', 'available_cans', 'target_quantity', 'status', 'message'],
  ORDERS_PROCESSED: ['timestamp', 'shopify_order_id', 'shopify_order_name', 'webhook_id', 'status', 'note']
};

async function sheetsClient() {
  let raw;
  if (config.googleCredentialsBase64) {
    raw = Buffer.from(config.googleCredentialsBase64, 'base64').toString('utf8');
  } else if (config.googleCredentialsPath) {
    raw = fs.readFileSync(config.googleCredentialsPath, 'utf8');
  } else {
    throw new Error('Missing Google service account JSON');
  }

  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(raw),
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });
  return google.sheets({ version: 'v4', auth: await auth.getClient() });
}

export async function readRows(sheetName) {
  const sheets = await sheetsClient();
  const response = await sheets.spreadsheets.values.get({ spreadsheetId: config.googleSheetId, range: `'${sheetName}'!A:Z` });
  const values = response.data.values || [];
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h).trim());
  return values.slice(1).filter(row => row.some(cell => String(cell || '').trim())).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] ?? ''; });
    return obj;
  });
}

export async function appendRows(sheetName, headers, objects) {
  if (!objects.length) return;
  const sheets = await sheetsClient();
  await sheets.spreadsheets.values.append({
    spreadsheetId: config.googleSheetId,
    range: `'${sheetName}'!A:Z`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: objects.map(obj => headers.map(h => obj[h] ?? '')) }
  });
}
