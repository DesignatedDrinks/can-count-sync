function yes(value) {
  const v = String(value ?? '').trim().toLowerCase();
  return v !== 'false' && v !== '0' && v !== 'no';
}

function num(value, fallback = 0) {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

export function inferPackSize({ variantTitle = '', sku = '' }) {
  const text = `${variantTitle} ${sku}`.toLowerCase();
  if (text.includes('12-pack') || text.includes('12 pack') || text.includes('12pk')) return 12;
  if (text.includes('6-pack') || text.includes('6 pack') || text.includes('6pk')) return 6;
  if (text.includes('4-pack') || text.includes('4 pack') || text.includes('4pk')) return 4;
  if (text.includes('24-pack') || text.includes('24 pack') || text.includes('24pk')) return 24;
  if (text.includes('single')) return 1;
  return '';
}

export function proposeBaseSku({ productHandle = '', productTitle = '', sku = '' }) {
  const raw = String(sku || productHandle || productTitle || '').toUpperCase();
  return raw
    .replace('SINGLE', '')
    .replace('4PK', '')
    .replace('12PK', '')
    .replace('4-PACK', '')
    .replace('12-PACK', '')
    .replace(/[^A-Z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 60);
}

export function calculateBaseInventory(baseProducts, ledgerRows, defaultSafetyBuffer = 0) {
  const result = new Map();

  for (const row of baseProducts) {
    const baseSku = String(row.base_sku || '').trim();
    if (!baseSku || !yes(row.active)) continue;
    result.set(baseSku, {
      base_sku: baseSku,
      product_name: row.product_name || baseSku,
      opening_cans: num(row.opening_cans),
      ledger_change: 0,
      safety_buffer_cans: num(row.safety_buffer_cans, defaultSafetyBuffer),
      available_cans: 0
    });
  }

  for (const row of ledgerRows) {
    const baseSku = String(row.base_sku || '').trim();
    if (!result.has(baseSku)) continue;
    result.get(baseSku).ledger_change += num(row.change_cans);
  }

  for (const item of result.values()) {
    item.available_cans = Math.max(0, Math.floor(item.opening_cans + item.ledger_change - item.safety_buffer_cans));
  }

  return result;
}

export function calculateVariantTargets(shopifyVariants, baseInventory) {
  const targets = [];
  const skipped = [];

  for (const variant of shopifyVariants) {
    if (!yes(variant.active) || !yes(variant.tracked)) continue;

    const baseSku = String(variant.base_sku || '').trim();
    const packSize = num(variant.pack_size);
    const inventoryItemId = String(variant.inventory_item_id || '').trim();

    if (!baseSku || !packSize || !inventoryItemId) {
      skipped.push({ variant, reason: 'Missing base_sku, pack_size, or inventory_item_id' });
      continue;
    }

    const base = baseInventory.get(baseSku);
    if (!base) {
      skipped.push({ variant, reason: `No active Base Products row for ${baseSku}` });
      continue;
    }

    targets.push({
      ...variant,
      base_sku: baseSku,
      pack_size: packSize,
      available_cans: base.available_cans,
      target_quantity: Math.max(0, Math.floor(base.available_cans / packSize))
    });
  }

  return { targets, skipped };
}

export function buildDeductionRowsFromOrder(order, variantMap, timestamp, webhookId = '') {
  const rows = [];

  for (const line of order.line_items || []) {
    const legacy = String(line.variant_id || '').trim();
    const gid = legacy ? `gid://shopify/ProductVariant/${legacy}` : '';
    const variant = variantMap.get(gid) || variantMap.get(legacy);
    if (!variant) continue;

    const quantity = num(line.quantity);
    const packSize = num(variant.pack_size);
    if (!quantity || !packSize) continue;

    rows.push({
      timestamp,
      base_sku: variant.base_sku,
      change_cans: -(quantity * packSize),
      reason: 'order',
      reference_type: 'shopify_order',
      reference_id: String(order.id || order.name || ''),
      note: `${order.name || 'Shopify order'} - ${quantity} x ${variant.variant_title || line.title} x ${packSize} cans`,
      source: webhookId ? `webhook:${webhookId}` : 'shopify_order'
    });
  }

  return rows;
}
