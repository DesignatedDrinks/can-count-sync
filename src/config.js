import dotenv from 'dotenv';

dotenv.config();

function required(name) {
  const value = process.env[name];
  if (!value || !String(value).trim()) throw new Error(`Missing required environment variable: ${name}`);
  return String(value).trim();
}

export function normalizeShop(shop) {
  return String(shop || '')
    .trim()
    .replace(/^https?:\/\//, '')
    .replace(/\/$/, '')
    .replace('.myshopify.com', '');
}

export const config = {
  shop: normalizeShop(required('SHOPIFY_SHOP')),
  clientId: required('SHOPIFY_CLIENT_ID'),
  clientSecret: required('SHOPIFY_CLIENT_SECRET'),
  apiVersion: process.env.SHOPIFY_API_VERSION || '2026-04',
  locationId: required('SHOPIFY_LOCATION_ID'),
  googleSheetId: required('GOOGLE_SHEET_ID'),
  googleCredentialsPath: process.env.GOOGLE_APPLICATION_CREDENTIALS || '',
  googleCredentialsBase64: process.env.GOOGLE_SERVICE_ACCOUNT_JSON_BASE64 || '',
  port: Number(process.env.PORT || 3000),
  webhookSecret: process.env.SHOPIFY_WEBHOOK_SECRET || process.env.SHOPIFY_CLIENT_SECRET || '',
  defaultSafetyBufferCans: Number(process.env.DEFAULT_SAFETY_BUFFER_CANS || 0)
};

export function nowIso() {
  return new Date().toISOString();
}

export function variantGid(value) {
  const s = String(value || '').trim();
  if (!s) return '';
  if (s.startsWith('gid://shopify/ProductVariant/')) return s;
  if (/^\d+$/.test(s)) return `gid://shopify/ProductVariant/${s}`;
  return s;
}

export function legacyId(gid) {
  const match = String(gid || '').match(/\/(\d+)$/);
  return match ? match[1] : '';
}
