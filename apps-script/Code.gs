const CCS = {
  tabs: { mapping: 'Pack Mapping', log: 'Sync Log', orders: 'Orders Processed' },
  headers: {
    mapping: ['sync_group','product_title','variant_title','sku','variant_gid','inventory_item_id','pack_size','is_master','active','shopify_available','target_available','tracked','last_seen_at','notes'],
    log: ['timestamp','mode','sync_group','product_title','variant_title','pack_size','shopify_available','target_available','status','message'],
    orders: ['timestamp','shopify_order_id','shopify_order_name','status','note']
  }
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Can Count Sync')
    .addItem('Setup Simple Tabs', 'ccsSetupTabs')
    .addSeparator()
    .addItem('Import / Refresh ACTIVE Shopify Products', 'ccsImportOrRefreshMapping')
    .addItem('Preview Safe Pack Sync', 'ccsPreviewPackSync')
    .addItem('Sync Safe Pack Inventory to Shopify', 'ccsLivePackSync')
    .addSeparator()
    .addItem('Process Recent Orders + Sync', 'ccsProcessRecentOrders')
    .addItem('Install 10 Minute Order Polling', 'ccsInstallOrderPolling')
    .addToUi();
}

function ccsSetupTabs() {
  ccsSetupTabsSilent();
  SpreadsheetApp.getUi().alert('Ready. Shopify Single Can inventory is the source of truth. Only ACTIVE Shopify products are imported. Only rows with active=TRUE are calculated or synced.');
}

function ccsSetupTabsSilent() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(CCS.tabs).forEach(k => {
    let sh = ss.getSheetByName(CCS.tabs[k]);
    if (!sh) sh = ss.insertSheet(CCS.tabs[k]);
    sh.getRange(1, 1, 1, CCS.headers[k].length).setValues([CCS.headers[k]]);
    sh.setFrozenRows(1);
  });
}

function ccsProps() { return PropertiesService.getScriptProperties(); }
function ccsOptional(name) { const v = ccsProps().getProperty(name); return v === null ? '' : String(v).trim(); }
function ccsGet(name, fallback) { const v = ccsOptional(name); if (v) return v; if (fallback !== undefined) return fallback; throw new Error('Missing script property: ' + name); }
function ccsShop() { return ccsGet('SHOPIFY_SHOP').replace('https://','').replace('http://','').replace('.myshopify.com','').replace('/','').trim(); }

function ccsToken() {
  const direct = ccsOptional('SHOPIFY_ADMIN_API_ACCESS_TOKEN');
  if (direct) return direct;
  const cache = CacheService.getScriptCache();
  const cached = cache.get('shopify_admin_token');
  if (cached) return cached;

  const url = 'https://' + ccsShop() + '.myshopify.com/admin/oauth/' + 'access' + '_token';
  const payload = { grant_type: 'client' + '_credentials', client_id: ccsGet('SHOPIFY_CLIENT_ID') };
  payload['client' + '_secret'] = ccsGet('SHOPIFY_CLIENT_SECRET');
  const res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
  const text = res.getContentText();
  const body = JSON.parse(text);
  if (res.getResponseCode() >= 300 || !body['access' + '_token']) throw new Error('Shopify token error: ' + text);
  const token = body['access' + '_token'];
  cache.put('shopify_admin_token', token, Math.min(Number(body.expires_in || 3600) - 60, 21600));
  return token;
}

function ccsGraphql(query, variables) {
  const url = 'https://' + ccsShop() + '.myshopify.com/admin/api/' + ccsGet('SHOPIFY_API_VERSION', '2026-04') + '/graphql.json';
  const headers = { 'Content-Type': 'application/json' };
  headers['X-Shopify-Access-Token'] = ccsToken();
  const res = UrlFetchApp.fetch(url, { method: 'post', headers, payload: JSON.stringify({ query, variables: variables || {} }), muteHttpExceptions: true });
  const text = res.getContentText();
  const body = JSON.parse(text);
  if (res.getResponseCode() >= 300 || body.errors) throw new Error('Shopify GraphQL error: ' + text);
  return body.data;
}

function ccsYes(v) { const s = String(v === undefined ? '' : v).trim().toLowerCase(); return s === 'true' || s === 'yes' || s === '1' || s === 'y'; }
function ccsNumber(v) { const n = Number(v); return isNaN(n) ? 0 : n; }

function ccsRead(tabName) {
  const sh = SpreadsheetApp.getActive().getSheetByName(tabName);
  if (!sh) return [];
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1).filter(r => r.some(c => String(c).trim() !== '')).map(row => {
    const o = {};
    headers.forEach((h, i) => o[h] = row[i]);
    return o;
  });
}

function ccsWriteRows(tabName, headers, rows) {
  let sh = SpreadsheetApp.getActive().getSheetByName(tabName);
  if (!sh) sh = SpreadsheetApp.getActive().insertSheet(tabName);
  sh.clearContents();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows.map(o => headers.map(h => o[h] === undefined ? '' : o[h])));
  sh.setFrozenRows(1);
}

function ccsAppend(tabName, headers, rows) {
  if (!rows.length) return;
  let sh = SpreadsheetApp.getActive().getSheetByName(tabName);
  if (!sh) { sh = SpreadsheetApp.getActive().insertSheet(tabName); sh.getRange(1, 1, 1, headers.length).setValues([headers]); }
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows.map(o => headers.map(h => o[h] === undefined ? '' : o[h])));
}

function ccsAvailableFromNode(v) {
  const qs = v.inventoryItem && v.inventoryItem.inventoryLevel && v.inventoryItem.inventoryLevel.quantities ? v.inventoryItem.inventoryLevel.quantities : [];
  const q = qs.find(x => x.name === 'available');
  return q ? Number(q.quantity || 0) : 0;
}

function ccsInferPackSize(variantTitle, sku, productTitle) {
  const text = (String(productTitle || '') + ' ' + String(variantTitle || '') + ' ' + String(sku || '')).toLowerCase();
  const pack = text.match(/\b([0-9]+)\s*(?:x\s*)?-?\s*pack\b/);
  if (pack) return Number(pack[1]);
  const pk = text.match(/\b([0-9]+)\s*pk\b/);
  if (pk) return Number(pk[1]);
  if (/\bsingle\b/.test(text) || /\bcan\b/.test(text)) return 1;
  if (/\b(237|250|330|341|355|440|473|500)\s*ml\b/.test(text) || /\b(237|250|330|341|355|440|473|500)ml\b/.test(text)) return 1;
  return '';
}

function ccsProposeGroup(productTitle, sku) {
  return String(productTitle || sku || '').toUpperCase()
    .replace(/\(NON-ALCOHOLIC\)/g, '')
    .replace(/\(NON ALCOHOLIC\)/g, '')
    .replace(/[^A-Z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 60);
}

function ccsFetchShopifyVariants() {
  const query = 'query ProductVariants($cursor: String, $locationId: ID!) { productVariants(first: 100, after: $cursor) { pageInfo { hasNextPage endCursor } nodes { id title sku product { title status } inventoryItem { id sku tracked inventoryLevel(locationId: $locationId) { quantities(names: ["available"]) { name quantity } } } } } }';
  let cursor = null;
  const rows = [];
  do {
    const data = ccsGraphql(query, { cursor, locationId: ccsGet('SHOPIFY_LOCATION_ID') });
    data.productVariants.nodes.forEach(v => {
      if (!v.product || String(v.product.status) !== 'ACTIVE') return;
      rows.push({
        product_title: v.product.title || '',
        variant_title: v.title || '',
        sku: v.sku || (v.inventoryItem && v.inventoryItem.sku) || '',
        variant_gid: v.id || '',
        inventory_item_id: v.inventoryItem ? v.inventoryItem.id : '',
        shopify_available: ccsAvailableFromNode(v),
        tracked: v.inventoryItem && v.inventoryItem.tracked ? 'TRUE' : 'FALSE'
      });
    });
    cursor = data.productVariants.pageInfo.hasNextPage ? data.productVariants.pageInfo.endCursor : null;
  } while (cursor);
  return rows;
}

function ccsImportOrRefreshMapping() {
  const rows = ccsRefreshMappingRows(true);
  SpreadsheetApp.getUi().alert('Pack Mapping refreshed with ' + rows.length + ' variants from ACTIVE Shopify products only. Rows with active=FALSE do not calculate or sync inventory.');
}

function ccsRefreshMappingRows(writeSheet) {
  ccsSetupTabsSilent();
  const existing = {};
  ccsRead(CCS.tabs.mapping).forEach(r => { if (r.variant_gid) existing[String(r.variant_gid)] = r; });
  const rows = ccsFetchShopifyVariants().map(v => {
    const old = existing[String(v.variant_gid)] || {};
    const packSize = String(old.pack_size || '').trim() || ccsInferPackSize(v.variant_title, v.sku, v.product_title);
    const isActive = ccsYes(old.active);
    return {
      sync_group: old.sync_group || ccsProposeGroup(v.product_title, v.sku),
      product_title: v.product_title,
      variant_title: v.variant_title,
      sku: v.sku,
      variant_gid: v.variant_gid,
      inventory_item_id: v.inventory_item_id,
      pack_size: packSize,
      is_master: (ccsYes(old.is_master) || String(packSize) === '1') ? 'TRUE' : 'FALSE',
      active: isActive ? 'TRUE' : 'FALSE',
      shopify_available: isActive ? v.shopify_available : '',
      target_available: isActive ? (old.target_available || '') : '',
      tracked: v.tracked,
      last_seen_at: new Date(),
      notes: old.notes || ''
    };
  });
  rows.sort((a, b) => String(a.product_title).localeCompare(String(b.product_title)) || ccsNumber(a.pack_size) - ccsNumber(b.pack_size));
  if (writeSheet) ccsWriteRows(CCS.tabs.mapping, CCS.headers.mapping, rows);
  return rows;
}

function ccsPreviewPackSync() { ccsRunPackSync('preview', false); }
function ccsLivePackSync() { ccsRunPackSync('live', false); }

function ccsFindSingleSource(groupRows) {
  return groupRows.find(r => ccsYes(r.active) && ccsYes(r.tracked) && ccsNumber(r.pack_size) === 1) || null;
}

function ccsPackTarget(totalCans, packSize) {
  if (packSize === 1) return totalCans;
  if (packSize >= 12) {
    if (totalCans < 20) return 0;
    return Math.max(1, Math.floor(totalCans / 24));
  }
  if (totalCans < packSize * 2) return 0;
  if (totalCans < 24) return Math.floor((totalCans - packSize) / packSize);
  return Math.max(4, Math.floor(totalCans / (packSize * 2)));
}

function ccsRunPackSync(mode, silent) {
  const rows = ccsRefreshMappingRows(false);
  const groups = {};
  rows.forEach(r => {
    if (!ccsYes(r.active)) return;
    const g = String(r.sync_group || '').trim();
    if (!g) return;
    if (!groups[g]) groups[g] = [];
    groups[g].push(r);
  });

  const logRows = [];
  Object.keys(groups).forEach(group => {
    const groupRows = groups[group];
    const source = ccsFindSingleSource(groupRows);
    if (!source) {
      logRows.push({ timestamp: new Date(), mode, sync_group: group, status: 'skipped', message: 'No active tracked Single Can source row found' });
      return;
    }

    const totalCans = ccsNumber(source.shopify_available);
    groupRows.forEach(r => {
      const pack = ccsNumber(r.pack_size);
      if (!pack || !r.inventory_item_id || !ccsYes(r.tracked)) {
        logRows.push({ timestamp: new Date(), mode, sync_group: group, product_title: r.product_title, variant_title: r.variant_title, pack_size: r.pack_size, shopify_available: r.shopify_available, status: 'skipped', message: 'Missing pack_size, inventory_item_id, or tracked=FALSE' });
        return;
      }

      const current = ccsNumber(r.shopify_available);
      const target = ccsPackTarget(totalCans, pack);
      r.target_available = target;
      let status = 'preview';
      let message = 'No Shopify update made';

      if (mode === 'live') {
        try {
          if (current !== target) {
            ccsSetInventory(r.inventory_item_id, target, current);
            r.shopify_available = target;
            status = 'success';
            message = 'Updated Shopify inventory';
          } else {
            status = 'unchanged';
            message = 'Already correct';
          }
        } catch (err) {
          status = 'error';
          message = err.message;
        }
      }

      logRows.push({ timestamp: new Date(), mode, sync_group: group, product_title: r.product_title, variant_title: r.variant_title, pack_size: pack, shopify_available: current, target_available: target, status, message });
    });
  });

  ccsWriteRows(CCS.tabs.mapping, CCS.headers.mapping, rows);
  ccsAppend(CCS.tabs.log, CCS.headers.log, logRows);
  if (!silent) SpreadsheetApp.getUi().alert(mode === 'live' ? 'Live safe pack sync complete. Check Sync Log.' : 'Preview complete. Check target_available and Sync Log.');
}

function ccsSetInventory(inventoryItemId, quantity, changeFromQuantity) {
  const mutation = 'mutation inventorySetQuantities($input: InventorySetQuantitiesInput!, $idempotencyKey: String!) { inventorySetQuantities(input: $input) @idempotent(key: $idempotencyKey) { userErrors { code field message } } }';
  const idempotencyKey = Utilities.getUuid();
  const input = {
    name: 'available',
    reason: 'correction',
    referenceDocumentUri: 'can-count-sync://' + idempotencyKey,
    quantities: [{ inventoryItemId, locationId: ccsGet('SHOPIFY_LOCATION_ID'), quantity: Number(quantity), changeFromQuantity: Number(changeFromQuantity) }]
  };
  const data = ccsGraphql(mutation, { input, idempotencyKey });
  const errors = data.inventorySetQuantities.userErrors || [];
  if (errors.length) throw new Error(JSON.stringify(errors));
}

function ccsProcessRecentOrders() {
  const rows = ccsRefreshMappingRows(false);
  const activeByVariant = {};
  const groups = {};
  rows.forEach(r => {
    if (!ccsYes(r.active)) return;
    activeByVariant[String(r.variant_gid)] = r;
    const g = String(r.sync_group || '').trim();
    if (!groups[g]) groups[g] = [];
    groups[g].push(r);
  });

  const processed = {};
  ccsRead(CCS.tabs.orders).forEach(r => processed[String(r.shopify_order_id)] = true);
  const since = new Date(Date.now() - Number(ccsGet('ORDER_LOOKBACK_HOURS', '24')) * 60 * 60 * 1000).toISOString();
  const query = 'query Orders($q: String!) { orders(first: 50, query: $q, sortKey: CREATED_AT, reverse: true) { nodes { id name lineItems(first: 100) { nodes { quantity variant { id title } } } } } }';
  const data = ccsGraphql(query, { q: 'created_at:>=' + since });

  const adjustments = {};
  const orderRows = [];
  data.orders.nodes.forEach(order => {
    if (processed[String(order.id)]) return;
    let touched = false;
    order.lineItems.nodes.forEach(line => {
      const sold = activeByVariant[String(line.variant ? line.variant.id : '')];
      if (!sold) return;
      const pack = ccsNumber(sold.pack_size);
      if (pack <= 1) return;
      const source = ccsFindSingleSource(groups[String(sold.sync_group || '').trim()] || []);
      if (!source) return;
      const cans = Number(line.quantity || 0) * pack;
      if (!adjustments[source.variant_gid]) adjustments[source.variant_gid] = { row: source, qty: 0, notes: [] };
      adjustments[source.variant_gid].qty += cans;
      adjustments[source.variant_gid].notes.push(order.name + ' sold ' + line.quantity + ' x ' + sold.variant_title + ' = ' + cans + ' cans');
      touched = true;
    });
    if (touched) orderRows.push({ timestamp: new Date(), shopify_order_id: order.id, shopify_order_name: order.name, status: 'processed', note: 'Pack sale adjusted Single Can source inventory' });
  });

  const logRows = [];
  Object.keys(adjustments).forEach(k => {
    const adj = adjustments[k];
    const current = ccsNumber(adj.row.shopify_available);
    const target = Math.max(0, current - adj.qty);
    try {
      ccsSetInventory(adj.row.inventory_item_id, target, current);
      logRows.push({ timestamp: new Date(), mode: 'order_adjust', sync_group: adj.row.sync_group, product_title: adj.row.product_title, variant_title: adj.row.variant_title, pack_size: adj.row.pack_size, shopify_available: current, target_available: target, status: 'success', message: adj.notes.join('; ') });
    } catch (err) {
      logRows.push({ timestamp: new Date(), mode: 'order_adjust', sync_group: adj.row.sync_group, product_title: adj.row.product_title, variant_title: adj.row.variant_title, pack_size: adj.row.pack_size, shopify_available: current, target_available: target, status: 'error', message: err.message });
    }
  });

  ccsAppend(CCS.tabs.orders, CCS.headers.orders, orderRows);
  ccsAppend(CCS.tabs.log, CCS.headers.log, logRows);
  ccsRunPackSync('live', true);
}

function ccsInstallOrderPolling() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'ccsProcessRecentOrders') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('ccsProcessRecentOrders').timeBased().everyMinutes(10).create();
  SpreadsheetApp.getUi().alert('10 minute order polling is installed. Pack sales deduct cans from the Single Can source, then safe pack inventory syncs.');
}
