const CCS = {
  tabs: {
    base: 'Base Products',
    variants: 'Shopify Variants',
    ledger: 'Can Ledger',
    needs: 'Needs Mapping',
    syncLog: 'Sync Log',
    orders: 'Orders Processed'
  },
  headers: {
    base: ['base_sku', 'product_name', 'case_size', 'opening_cans', 'safety_buffer_cans', 'active'],
    variants: ['base_sku', 'product_title', 'variant_title', 'sku', 'price', 'variant_gid', 'variant_legacy_id', 'inventory_item_id', 'pack_size', 'active', 'tracked', 'last_seen_at'],
    ledger: ['timestamp', 'base_sku', 'change_cans', 'reason', 'reference_type', 'reference_id', 'note', 'source'],
    needs: ['status', 'product_title', 'variant_title', 'sku', 'price', 'proposed_base_sku', 'proposed_pack_size', 'variant_gid', 'variant_legacy_id', 'inventory_item_id', 'tracked', 'last_seen_at', 'note'],
    syncLog: ['timestamp', 'mode', 'base_sku', 'product_title', 'variant_title', 'pack_size', 'available_cans', 'target_quantity', 'status', 'message'],
    orders: ['timestamp', 'shopify_order_id', 'shopify_order_name', 'status', 'note']
  }
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Can Count Sync')
    .addItem('Setup / Repair Sheet Tabs', 'ccsSetupTabs')
    .addSeparator()
    .addItem('Import Shopify Variants', 'ccsImportShopifyVariants')
    .addItem('Promote Ready Mapping Rows', 'ccsPromoteMappings')
    .addSeparator()
    .addItem('Preview Inventory Sync', 'ccsPreviewInventorySync')
    .addItem('Sync Inventory to Shopify', 'ccsSyncInventoryToShopify')
    .addSeparator()
    .addItem('Process Recent Orders', 'ccsProcessRecentOrders')
    .addItem('Install 10 Minute Order Polling', 'ccsInstallOrderPolling')
    .addToUi();
}

function ccsSetupTabs() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(CCS.tabs).forEach(key => {
    let sh = ss.getSheetByName(CCS.tabs[key]);
    if (!sh) sh = ss.insertSheet(CCS.tabs[key]);
    const headers = CCS.headers[key];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  });
  SpreadsheetApp.getUi().alert('Can Count Sync tabs are ready.');
}

function ccsProps() {
  return PropertiesService.getScriptProperties();
}

function ccsGet(name, fallback) {
  const value = ccsProps().getProperty(name);
  if (value === null || value === '') {
    if (fallback !== undefined) return fallback;
    throw new Error('Missing script property: ' + name);
  }
  return value;
}

function ccsShop() {
  return ccsGet('SHOPIFY_SHOP').replace('https://', '').replace('http://', '').replace('.myshopify.com', '').replace('/', '').trim();
}

function ccsToken() {
  return ccsGet('SHOPIFY_ADMIN_API_ACCESS_TOKEN');
}

function ccsGraphql(query, variables) {
  const version = ccsGet('SHOPIFY_API_VERSION', '2026-04');
  const url = 'https://' + ccsShop() + '.myshopify.com/admin/api/' + version + '/graphql.json';
  const headers = { 'Content-Type': 'application/json' };
  headers['X-Shopify-Access-Token'] = ccsToken();
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: headers,
    payload: JSON.stringify({ query: query, variables: variables || {} }),
    muteHttpExceptions: true
  });
  const text = res.getContentText();
  let body;
  try {
    body = JSON.parse(text);
  } catch (err) {
    throw new Error('Shopify returned non-JSON. Check SHOPIFY_SHOP and SHOPIFY_ADMIN_API_ACCESS_TOKEN. First 300 chars: ' + text.slice(0, 300));
  }
  if (res.getResponseCode() >= 300 || body.errors) throw new Error('Shopify GraphQL error: ' + text);
  return body.data;
}

function ccsRead(tabName) {
  const sh = SpreadsheetApp.getActive().getSheetByName(tabName);
  if (!sh) return [];
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1).filter(r => r.some(c => String(c).trim() !== '')).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function ccsAppend(tabName, headers, rows) {
  if (!rows.length) return;
  let sh = SpreadsheetApp.getActive().getSheetByName(tabName);
  if (!sh) {
    sh = SpreadsheetApp.getActive().insertSheet(tabName);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  const values = rows.map(obj => headers.map(h => obj[h] === undefined ? '' : obj[h]));
  sh.getRange(sh.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
}

function ccsYes(v) {
  const s = String(v === undefined ? '' : v).trim().toLowerCase();
  return s !== 'false' && s !== '0' && s !== 'no';
}

function ccsNumber(v) {
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function ccsLegacyId(gid) {
  const m = String(gid || '').match(/\/(\d+)$/);
  return m ? m[1] : '';
}

function ccsInferPackSize(title, sku) {
  const text = (String(title || '') + ' ' + String(sku || '')).toLowerCase();
  if (text.indexOf('12-pack') >= 0 || text.indexOf('12 pack') >= 0 || text.indexOf('12pk') >= 0) return 12;
  if (text.indexOf('6-pack') >= 0 || text.indexOf('6 pack') >= 0 || text.indexOf('6pk') >= 0) return 6;
  if (text.indexOf('4-pack') >= 0 || text.indexOf('4 pack') >= 0 || text.indexOf('4pk') >= 0) return 4;
  if (text.indexOf('24-pack') >= 0 || text.indexOf('24 pack') >= 0 || text.indexOf('24pk') >= 0) return 24;
  if (text.indexOf('single') >= 0) return 1;
  return '';
}

function ccsProposeBaseSku(productTitle, sku) {
  return String(sku || productTitle || '').toUpperCase()
    .replace('SINGLE', '')
    .replace('4PK', '')
    .replace('12PK', '')
    .replace('4-PACK', '')
    .replace('12-PACK', '')
    .replace(/[^A-Z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 60);
}

function ccsImportShopifyVariants() {
  ccsSetupTabs();
  const existing = ccsRead(CCS.tabs.variants).concat(ccsRead(CCS.tabs.needs));
  const known = {};
  existing.forEach(r => {
    if (r.variant_gid) known[String(r.variant_gid)] = true;
    if (r.variant_legacy_id) known[String(r.variant_legacy_id)] = true;
  });

  const query = 'query ProductVariants($cursor: String) { productVariants(first: 100, after: $cursor) { pageInfo { hasNextPage endCursor } nodes { id legacyResourceId title sku price product { title handle } inventoryItem { id sku tracked } } } }';
  let cursor = null;
  const rows = [];
  do {
    const data = ccsGraphql(query, { cursor: cursor });
    data.productVariants.nodes.forEach(v => {
      const legacy = ccsLegacyId(v.id) || v.legacyResourceId || '';
      if (known[v.id] || known[legacy]) return;
      const sku = v.sku || (v.inventoryItem && v.inventoryItem.sku) || '';
      rows.push({
        status: 'needs_mapping',
        product_title: v.product ? v.product.title : '',
        variant_title: v.title || '',
        sku: sku,
        price: v.price || '',
        proposed_base_sku: ccsProposeBaseSku(v.product ? v.product.title : '', sku),
        proposed_pack_size: ccsInferPackSize(v.title, sku),
        variant_gid: v.id || '',
        variant_legacy_id: legacy,
        inventory_item_id: v.inventoryItem ? v.inventoryItem.id : '',
        tracked: v.inventoryItem && v.inventoryItem.tracked ? 'TRUE' : 'FALSE',
        last_seen_at: new Date(),
        note: 'Confirm base SKU and pack size, then set status to ready_to_promote.'
      });
    });
    cursor = data.productVariants.pageInfo.hasNextPage ? data.productVariants.pageInfo.endCursor : null;
  } while (cursor);

  ccsAppend(CCS.tabs.needs, CCS.headers.needs, rows);
  SpreadsheetApp.getUi().alert('Imported ' + rows.length + ' new variants into Needs Mapping.');
}

function ccsPromoteMappings() {
  const ready = ccsRead(CCS.tabs.needs).filter(r => String(r.status).trim().toLowerCase() === 'ready_to_promote');
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
    last_seen_at: new Date()
  })).filter(r => r.base_sku && r.pack_size && r.inventory_item_id);
  ccsAppend(CCS.tabs.variants, CCS.headers.variants, rows);
  SpreadsheetApp.getUi().alert('Promoted ' + rows.length + ' rows.');
}

function ccsBaseInventory() {
  const baseRows = ccsRead(CCS.tabs.base);
  const ledger = ccsRead(CCS.tabs.ledger);
  const map = {};
  baseRows.forEach(r => {
    const base = String(r.base_sku || '').trim();
    if (!base || !ccsYes(r.active)) return;
    map[base] = {
      base_sku: base,
      product_name: r.product_name || base,
      opening_cans: ccsNumber(r.opening_cans),
      ledger_change: 0,
      safety_buffer_cans: ccsNumber(r.safety_buffer_cans),
      available_cans: 0
    };
  });
  ledger.forEach(r => {
    const base = String(r.base_sku || '').trim();
    if (map[base]) map[base].ledger_change += ccsNumber(r.change_cans);
  });
  Object.keys(map).forEach(base => {
    const item = map[base];
    item.available_cans = Math.max(0, Math.floor(item.opening_cans + item.ledger_change - item.safety_buffer_cans));
  });
  return map;
}

function ccsTargets() {
  const inv = ccsBaseInventory();
  return ccsRead(CCS.tabs.variants).filter(v => ccsYes(v.active) && ccsYes(v.tracked)).map(v => {
    const base = inv[String(v.base_sku || '').trim()];
    const pack = ccsNumber(v.pack_size);
    if (!base || !pack || !v.inventory_item_id) return null;
    return Object.assign({}, v, {
      available_cans: base.available_cans,
      target_quantity: Math.floor(base.available_cans / pack)
    });
  }).filter(Boolean);
}

function ccsPreviewInventorySync() {
  const targets = ccsTargets();
  const rows = targets.map(t => ({
    timestamp: new Date(), mode: 'preview', base_sku: t.base_sku, product_title: t.product_title, variant_title: t.variant_title, pack_size: t.pack_size, available_cans: t.available_cans, target_quantity: t.target_quantity, status: 'preview', message: 'No Shopify update made'
  }));
  ccsAppend(CCS.tabs.syncLog, CCS.headers.syncLog, rows);
  SpreadsheetApp.getUi().alert('Preview complete. Rows written to Sync Log: ' + rows.length);
}

function ccsSetInventory(t) {
  const mutation = 'mutation InventorySet($input: InventorySetQuantitiesInput!) { inventorySetQuantities(input: $input) { userErrors { field message code } } }';
  const input = {
    name: 'available',
    reason: 'can-count-sync',
    referenceDocumentUri: 'can-count-sync://' + new Date().getTime(),
    quantities: [{ inventoryItemId: t.inventory_item_id, locationId: ccsGet('SHOPIFY_LOCATION_ID'), quantity: Number(t.target_quantity) }]
  };
  const data = ccsGraphql(mutation, { input: input });
  const errors = data.inventorySetQuantities.userErrors || [];
  if (errors.length) throw new Error(JSON.stringify(errors));
}

function ccsSyncInventoryToShopify() {
  const targets = ccsTargets();
  const rows = [];
  targets.forEach(t => {
    try {
      ccsSetInventory(t);
      rows.push({ timestamp: new Date(), mode: 'live', base_sku: t.base_sku, product_title: t.product_title, variant_title: t.variant_title, pack_size: t.pack_size, available_cans: t.available_cans, target_quantity: t.target_quantity, status: 'success', message: 'Updated Shopify inventory' });
    } catch (err) {
      rows.push({ timestamp: new Date(), mode: 'live', base_sku: t.base_sku, product_title: t.product_title, variant_title: t.variant_title, pack_size: t.pack_size, available_cans: t.available_cans, target_quantity: t.target_quantity, status: 'error', message: err.message });
    }
  });
  ccsAppend(CCS.tabs.syncLog, CCS.headers.syncLog, rows);
  SpreadsheetApp.getUi().alert('Live sync complete. Check Sync Log.');
}

function ccsProcessRecentOrders() {
  const hours = Number(ccsGet('ORDER_LOOKBACK_HOURS', '24'));
  const since = new Date(Date.now() - hours * 60 * 60 * 1000).toISOString();
  const query = 'query Orders($q: String!) { orders(first: 50, query: $q, sortKey: CREATED_AT, reverse: true) { nodes { id name createdAt lineItems(first: 100) { nodes { quantity variant { id legacyResourceId title } } } } } }';
  const data = ccsGraphql(query, { q: 'created_at:>=' + since });

  const processed = {};
  ccsRead(CCS.tabs.orders).forEach(r => processed[String(r.shopify_order_id)] = true);

  const variants = {};
  ccsRead(CCS.tabs.variants).forEach(v => {
    variants[String(v.variant_gid)] = v;
    variants[String(v.variant_legacy_id)] = v;
  });

  const ledgerRows = [];
  const orderRows = [];
  data.orders.nodes.forEach(order => {
    if (processed[String(order.id)]) return;
    order.lineItems.nodes.forEach(line => {
      const v = line.variant || {};
      const match = variants[String(v.id)] || variants[String(v.legacyResourceId)];
      if (!match) return;
      const cans = Number(line.quantity) * Number(match.pack_size || 0);
      if (!cans) return;
      ledgerRows.push({ timestamp: new Date(), base_sku: match.base_sku, change_cans: -cans, reason: 'order', reference_type: 'shopify_order', reference_id: order.id, note: order.name + ' - ' + line.quantity + ' x ' + match.variant_title, source: 'order_polling' });
    });
    orderRows.push({ timestamp: new Date(), shopify_order_id: order.id, shopify_order_name: order.name, status: 'processed', note: 'Order polling processed' });
  });

  ccsAppend(CCS.tabs.ledger, CCS.headers.ledger, ledgerRows);
  ccsAppend(CCS.tabs.orders, CCS.headers.orders, orderRows);
  ccsSyncInventoryToShopify();
}

function ccsInstallOrderPolling() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'ccsProcessRecentOrders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('ccsProcessRecentOrders').timeBased().everyMinutes(10).create();
  SpreadsheetApp.getUi().alert('10 minute order polling is installed.');
}
