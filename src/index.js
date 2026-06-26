import assert from 'node:assert/strict';
import { config, nowIso, variantGid } from './config.js';
import { listLocations, fetchAllProductVariants, setInventoryQuantity, variantRowFromShopify } from './shopify.js';
import { SHEETS, HEADERS, readRows, appendRows } from './sheets.js';
import { inferPackSize, proposeBaseSku, calculateBaseInventory, calculateVariantTargets } from './calculator.js';

async function importVariants() {
  const existing = await readRows(SHEETS.SHOPIFY_VARIANTS);
  const needs = await readRows(SHEETS.NEEDS_MAPPING);
  const known = new Set([
    ...existing.map(r => r.variant_gid).filter(Boolean),
    ...existing.map(r => r.variant_legacy_id).filter(Boolean),
    ...needs.map(r => r.variant_gid).filter(Boolean),
    ...needs.map(r => r.variant_legacy_id).filter(Boolean)
  ]);

  const variants = await fetchAllProductVariants();
  const rows = [];
  for (const v of variants) {
    const row = variantRowFromShopify(v);
    if (known.has(row.variant_gid) || known.has(row.variant_legacy_id)) continue;
    rows.push({
      status: 'needs_mapping',
      ...row,
      proposed_base_sku: proposeBaseSku({ productHandle: v.product?.handle, productTitle: v.product?.title, sku: row.sku }),
      proposed_pack_size: inferPackSize({ variantTitle: row.variant_title, sku: row.sku }),
      last_seen_at: nowIso(),
      note: 'Confirm base SKU and pack size, then set status to ready_to_promote.'
    });
  }
  await appendRows(SHEETS.NEEDS_MAPPING, HEADERS.NEEDS_MAPPING, rows);
  console.log(`Fetched ${variants.length} Shopify variants. Added ${rows.length} new rows to Needs Mapping.`);
}

async function promoteMapping() {
  const needs = await readRows(SHEETS.NEEDS_MAPPING);
  const ready = needs.filter(r => String(r.status || '').trim().toLowerCase() === 'ready_to_promote');
  const rows = ready.map(r => ({
    base_sku: r.proposed_base_sku,
    product_title: r.product_title,
    variant_title: r.variant_title,
    sku: r.sku,
    price: r.price,
    variant_gid: r.variant_gid,
    variant_legacy_id: r.variant_legacy_id,
    inventory_item_id: r.inventory_item_id,
    pack_size: r.proposed_pack_size,
    active: 'TRUE',
    tracked: r.tracked || 'TRUE',
    last_seen_at: nowIso()
  })).filter(r => r.base_sku && r.pack_size && r.inventory_item_id);
  await appendRows(SHEETS.SHOPIFY_VARIANTS, HEADERS.SHOPIFY_VARIANTS, rows);
  console.log(`Promoted ${rows.length} mappings. Leave old Needs Mapping rows alone or mark them done manually.`);
}

async function buildTargets() {
  const baseProducts = await readRows(SHEETS.BASE_PRODUCTS);
  const ledger = await readRows(SHEETS.CAN_LEDGER);
  const variants = await readRows(SHEETS.SHOPIFY_VARIANTS);
  const baseInventory = calculateBaseInventory(baseProducts, ledger, config.defaultSafetyBufferCans);
  return calculateVariantTargets(variants, baseInventory);
}

async function sync({ live }) {
  const { targets, skipped } = await buildTargets();
  console.table(targets.map(t => ({ base_sku: t.base_sku, variant: t.variant_title, pack_size: t.pack_size, cans: t.available_cans, target: t.target_quantity })));

  const logRows = [];
  for (const target of targets) {
    if (!live) {
      logRows.push({ timestamp: nowIso(), mode: 'preview', ...target, status: 'preview', message: 'No Shopify update made' });
      continue;
    }
    try {
      await setInventoryQuantity(target, 'can-count-sync');
      logRows.push({ timestamp: nowIso(), mode: 'live', ...target, status: 'success', message: 'Updated Shopify inventory' });
    } catch (error) {
      logRows.push({ timestamp: nowIso(), mode: 'live', ...target, status: 'error', message: error.message });
    }
  }
  for (const item of skipped) {
    logRows.push({ timestamp: nowIso(), mode: live ? 'live' : 'preview', base_sku: item.variant?.base_sku || '', product_title: item.variant?.product_title || '', variant_title: item.variant?.variant_title || '', pack_size: item.variant?.pack_size || '', available_cans: '', target_quantity: '', status: 'skipped', message: item.reason });
  }
  await appendRows(SHEETS.SYNC_LOG, HEADERS.SYNC_LOG, logRows);
  console.log(`${live ? 'Live sync' : 'Preview'} complete. Targets: ${targets.length}. Skipped: ${skipped.length}.`);
}

async function test() {
  const base = [{ base_sku: 'BELL-PILSNER', opening_cans: 0, active: 'TRUE' }];
  const ledger = [{ base_sku: 'BELL-PILSNER', change_cans: 240 }, { base_sku: 'BELL-PILSNER', change_cans: -3 }];
  const variants = [
    { base_sku: 'BELL-PILSNER', variant_title: 'Single', pack_size: 1, inventory_item_id: '1', active: 'TRUE', tracked: 'TRUE' },
    { base_sku: 'BELL-PILSNER', variant_title: '4-Pack', pack_size: 4, inventory_item_id: '2', active: 'TRUE', tracked: 'TRUE' },
    { base_sku: 'BELL-PILSNER', variant_title: '12-Pack', pack_size: 12, inventory_item_id: '3', active: 'TRUE', tracked: 'TRUE' }
  ];
  const inv = calculateBaseInventory(base, ledger, 0);
  const { targets } = calculateVariantTargets(variants, inv);
  assert.deepEqual(targets.map(t => t.target_quantity), [237, 59, 19]);
  console.log('Test passed: 237 cans = Single 237 / 4-Pack 59 / 12-Pack 19');
}

const command = process.argv[2];
if (command === 'locations') {
  console.table(await listLocations());
} else if (command === 'import-variants') {
  await importVariants();
} else if (command === 'promote-mapping') {
  await promoteMapping();
} else if (command === 'sync-preview') {
  await sync({ live: false });
} else if (command === 'sync') {
  await sync({ live: true });
} else if (command === 'test') {
  await test();
} else {
  console.log('Use: locations | import-variants | promote-mapping | sync-preview | sync | test');
}
