import express from 'express';
import crypto from 'node:crypto';
import { config, nowIso, variantGid } from './config.js';
import { SHEETS, HEADERS, readRows, appendRows } from './sheets.js';
import { buildDeductionRowsFromOrder } from './calculator.js';
import './index.js';

function verify(rawBody, hmacHeader) {
  if (!config.webhookSecret || !hmacHeader) return false;
  const digest = crypto.createHmac('sha256', config.webhookSecret).update(rawBody, 'utf8').digest('base64');
  try {
    return crypto.timingSafeEqual(Buffer.from(digest), Buffer.from(hmacHeader));
  } catch {
    return false;
  }
}

function buildVariantMap(variants) {
  const map = new Map();
  for (const v of variants) {
    if (v.variant_gid) map.set(String(v.variant_gid).trim(), v);
    if (v.variant_legacy_id) map.set(String(v.variant_legacy_id).trim(), v);
    if (v.variant_legacy_id) map.set(variantGid(v.variant_legacy_id), v);
  }
  return map;
}

async function processOrder(order, webhookId) {
  const orderKey = String(order.id || order.name || '').trim();
  const processed = await readRows(SHEETS.ORDERS_PROCESSED);
  if (processed.some(r => String(r.shopify_order_id || '') === orderKey || String(r.webhook_id || '') === webhookId)) {
    return { status: 'already_processed', rows: [] };
  }

  const variants = await readRows(SHEETS.SHOPIFY_VARIANTS);
  const rows = buildDeductionRowsFromOrder(order, buildVariantMap(variants), nowIso(), webhookId);
  await appendRows(SHEETS.CAN_LEDGER, HEADERS.CAN_LEDGER, rows);
  await appendRows(SHEETS.ORDERS_PROCESSED, HEADERS.ORDERS_PROCESSED, [{
    timestamp: nowIso(),
    shopify_order_id: orderKey,
    shopify_order_name: order.name || '',
    webhook_id: webhookId,
    status: 'processed',
    note: `Added ${rows.length} ledger deduction rows`
  }]);
  return { status: 'processed', rows };
}

const app = express();

app.get('/health', (_req, res) => res.json({ ok: true, service: 'can-count-sync' }));

app.post('/webhooks/shopify', express.raw({ type: 'application/json' }), async (req, res) => {
  const raw = req.body.toString('utf8');
  const hmac = req.get('X-Shopify-Hmac-Sha256') || '';
  const topic = req.get('X-Shopify-Topic') || '';
  const webhookId = req.get('X-Shopify-Webhook-Id') || '';

  if (!verify(raw, hmac)) return res.status(401).send('Invalid webhook signature');

  const payload = JSON.parse(raw);
  if (topic === 'orders/create') await processOrder(payload, webhookId);
  if (topic === 'products/create' || topic === 'products/update') {
    console.log('Product webhook received. Run npm run import:variants to register new variants.');
  }

  res.status(200).send('ok');
});

app.listen(config.port, () => {
  console.log(`Can Count Sync server listening on port ${config.port}`);
});
