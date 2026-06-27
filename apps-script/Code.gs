const CCS = {
  tabs: {
    mapping: 'Pack Mapping',
    log: 'Sync Log',
    orders: 'Orders Processed'
  },
  headers: {
    mapping: ['sync_group', 'product_title', 'variant_title', 'sku', 'variant_gid', 'inventory_item_id', 'pack_size', 'is_master', 'active', 'shopify_available', 'target_available', 'tracked', 'last_seen_at', 'notes'],
    log: ['timestamp', 'mode', 'sync_group', 'product_title', 'variant_title', 'pack_size', 'shopify_available', 'target_available', 'status', 'message'],
    orders: ['timestamp', 'shopify_order_id', 'shopify_order_name', 'status', 'note']
  }
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Can Count Sync')
    .addItem('Setup Simple Tabs', 'ccsSetupTabs')
    .addSeparator()
    .addItem('Import / Refresh Shopify Variants', 'ccsImportOrRefreshMapping')
    .addItem('Preview Pack Sync', 'ccsPreviewPackSync')
    .addItem('Sync Pack Variants to Shopify', 'ccsLivePackSync')
    .addSeparator()
    .addItem('Process Recent Orders + Sync', 'ccsProcessRecentOrders')
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
  SpreadsheetApp.getUi().alert('Simple Can Count Sync tabs are ready. Use Pack Mapping only. Old Base Products / Ledger tabs can be ignored.');
}

function ccsProps() {
  return PropertiesService.getScriptProperties();
}

function ccsOptional(name) {
  const value = ccsProps().getProperty(name);
  return value === null ? '' : String(value).trim();
}

function ccsGet(name, fallback) {
  const value = ccsOptional(name);
  if (value === '') {
    if (fallback !== undefined) return fallback;
    throw new Error('Missing script property: ' + name);
  }
  return value;
}

function ccsShop() {
  return ccsGet('SHOPIFY_SHOP').replace('https://', '').replace('http://', '').replace('.myshopify.com', '').replace('/', '').trim();
}

function ccsToken() {
  const directToken = ccsOptional('SHOPIFY_ADMIN_API_ACCESS_TOKEN');
  if (directToken) return directToken;

  const cache = CacheService.getScriptCache();
  const cached = cache.get('shopify_admin_token');
  if (cached) return cached;

  const url = 'https://' + ccsShop() + '.myshopify.com/admin/oauth/' + 'access' + '_token';
  const payload = {
    grant_type: 'client' + '_credentials',
    client_id: ccsGet('SHOPIFY_CLIENT_ID')
  };
  payload['client' + '_secret'] = ccsGet('SHOPIFY_CLIENT_SECRET');

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const text = res.getContentText();
  let body;
  try {
    body = JSON.parse(text);
  } catch (err) {
    throw new Error('Shopify returned non-JSON during token request. Make sure the app is installed and SHOPIFY_SHOP is correct. First 300 chars: ' + text.slice(0, 300));
  }

  if (res.getResponseCode() >= 300 || !body['access' + '_token']) {
    throw new Error('Shopify token error: ' + text);
  }

  const token = body['access' + '_token'];
  cache.put('shopify_admin_token', token, Math.min(Number(body.expires_in || 3600) - 60, 21600));
  return token;
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
    throw new Error('Shopify returned non-JSON. Check app install and shop value. First 300 chars: ' + text.slice(0, 300));
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

function ccsWriteRows(tabName, headers, rows) {
  let sh = SpreadsheetApp.getActive().getSheetByName(tabName);
  if (!sh) sh = SpreadsheetApp.getActive().insertSheet(tabName);
  sh.clearContents();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    const values = rows.map(obj => headers.map(h => obj[h] === undefined ? '' : obj[h]));
    sh.getRange(2, 1, values.length, headers.length).setValues(values);
  }
  sh.setFrozenRows(1);
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
  return s === 'true' || s === 'yes' || s === '1' || s === 'y';
}

function ccsNumber(v) {
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function ccsAvailableFromNode(v) {
  const level = v.inventoryItem && v.inventoryItem.inventoryLevel;
  const quantities = level && level.quantities ? level.quantities : [];
  for (let i = 0; i < quantities.length; i++) {
    if (quantities[i].name === 'available') return Number(quantities[i].quantity || 0);
  }
  return 0;
}

function ccsInferPackSize(variantTitle, sku, productTitle) {
  const text = (String(productTitle || '') + ' ' + String(variantTitle || '') + ' ' + String(sku || '')).toLowerCase();
  if (text.indexOf('24-pack') >= 0 || text.indexOf('24 pack') >= 0 || text.indexOf('24pk') >= 0) return 24;
  if (text.indexOf('12-pack') >= 0 || text.indexOf('12 pack') >= 0 || text.indexOf('12pk') >= 0) return 12;
  if (text.indexOf('6-pack') >= 0 || text.indexOf('6 pack') >= 0 || text.indexOf('6pk') >= 0) return 6;
  if (text.indexOf('4-pack') >= 0 || text.indexOf('4 pack') >= 0 || text.indexOf('4pk') >= 0) return 4;
  if (text.indexOf('single') >= 0 || text.indexOf('1-pack') >= 0 || text.indexOf('1 pack') >= 0 || text.indexOf(' can') >= 0 || text.indexOf('can ') >= 0) return 1;
  if (/\b(250|330|341|355|473|500)\s*ml\b/.test(text)) return 1;
  if (/\b(250|330|341|355|473|500)ml\b/.test(text)) return 1;
  return '';
}

function ccsProposeGroup(productTitle, sku) {
  return String(productTitle || sku || '').toUpperCase()
    .replace(/\(NON-ALCOHOLIC\)/g, '')
    .replace(/[^A-Z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 60);
}

function ccsFetchShopifyVariants() {
  const query = 'query ProductVariants($cursor: String, $locationId: ID!) { productVariants(first: 100, after: $cursor) { pageInfo { hasNextPage endCursor } nodes { id title sku price product { title } inventoryItem { id sku tracked inventoryLevel(locationId: $locationId) { quantities(names: ["available"]) { name quantity } } } } } }';
  let cursor = null;
  const rows = [];
  do {
    const data = ccsGraphql(query, { cursor: cursor, locationId: ccsGet('SHOPIFY_LOCATION_ID') });
    data.productVariants.nodes.forEach(v => {
      const productTitle = v.product ? v.product.title : '';
      rows.push({
        product_title: productTitle,
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
  SpreadsheetApp.getUi().alert('Pack Mapping refreshed with ' + rows.length + ' Shopify variants. Set active=TRUE only for products you want Can Count Sync to manage.');
}

function ccsRefreshMappingRows(writeSheet) {
  ccsSetupTabsSilent();
  const existing = {};
  ccsRead(CCS.tabs.mapping).forEach(r => {
    if (r.variant_gid) existing[String(r.variant_gid)] = r;
  });

  const live = ccsFetchShopifyVariants();
  const rows = live.map(v => {
    const old = existing[String(v.variant_gid)] || {};
    const inferredPack = ccsInferPackSize(v.variant_title, v.sku, v.product_title);
    const oldPack = String(old.pack_size || '').trim();
    const packSize = oldPack || inferredPack;
    const masterValue = ccsYes(old.is_master) || String(packSize) === '1' ? 'TRUE' : 'FALSE';
    return {
      sync_group: old.sync_group || ccsProposeGroup(v.product_title, v.sku),
      product_title: v.product_title,
      variant_title: v.variant_title,
      sku: v.sku,
      variant_gid: v.variant_gid,
      inventory_item_id: v.inventory_item_id,
      pack_size: packSize,
      is_master: masterValue,
      active: old.active || 'FALSE',
      shopify_available: v.shopify_available,
      target_available: old.target_available || '',
      tracked: v.tracked,
      last_seen_at: new Date(),
      notes: old.notes || ''
    };
  });

  rows.sort((a, b) => {
    const p = String(a.product_title).localeCompare(String(b.product_title));
    if (p !== 0) return p;
    return ccsNumber(a.pack_size) - ccsNumber(b.pack_size);
  });

  if (writeSheet) ccsWriteRows(CCS.tabs.mapping, CCS.headers.mapping, rows);
  return rows;
}

function ccsSetupTabsSilent() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(CCS.tabs).forEach(key => {
    let sh = ss.getSheetByName(CCS.tabs[key]);
    if (!sh) {
      sh = ss.insertSheet(CCS.tabs[key]);
      sh.getRange(1, 1, 1, CCS.headers[key].length).setValues([CCS.headers[key]]);
      sh.setFrozenRows(1);
    }
  });
}

function ccsPreviewPackSync() {
  ccsRunPackSync('preview');
}

function ccsLivePackSync() {
  ccsRunPackSync('live');
}

function ccsFindMaster(groupRows) {
  const active = groupRows.filter(r => ccsYes(r.active) && ccsYes(r.tracked));
  let master = active.find(r => ccsYes(r.is_master));
  if (master) return master;
  master = active.find(r => ccsNumber(r.pack_size) === 1);
  if (master) return master;
  active.sort((a, b) => ccsNumber(a.pack_size) - ccsNumber(b.pack_size));
  return active[0] || null;
}

function ccsRunPackSync(mode) {
  const rows = ccsRefreshMappingRows(false);
  const groups = {};
  rows.forEach(r => {
    if (!ccsYes(r.active)) return;
    const group = String(r.sync_group || '').trim();
    if (!group) return;
    if (!groups[group]) groups[group] = [];
    groups[group].push(r);
  });

  const logRows = [];
  Object.keys(groups).forEach(group => {
    const groupRows = groups[group];
    const master = ccsFindMaster(groupRows);
    if (!master) {
      logRows.push({ timestamp: new Date(), mode: mode, sync_group: group, status: 'skipped', message: 'No active tracked master row found' });
      return;
    }

    const masterPack = ccsNumber(master.pack_size);
    const masterQty = ccsNumber(master.shopify_available);
    if (!masterPack) {
      logRows.push({ timestamp: new Date(), mode: mode, sync_group: group, status: 'skipped', message: 'Master pack_size missing' });
      return;
    }

    const totalCans = masterQty * masterPack;
    groupRows.forEach(r => {
      const pack = ccsNumber(r.pack_size);
      if (!pack || !r.inventory_item_id || !ccsYes(r.tracked)) {
        logRows.push({ timestamp: new Date(), mode: mode, sync_group: group, product_title: r.product_title, variant_title: r.variant_title, pack_size: r.pack_size, shopify_available: r.shopify_available, status: 'skipped', message: 'Missing pack_size, inventory_item_id, or tracked=FALSE' });
        return;
      }

      const target = Math.floor(totalCans / pack);
      r.target_available = target;
      let status = 'preview';
      let message = 'No Shopify update made';

      if (mode === 'live') {
        try {
          if (ccsNumber(r.shopify_available) !== target) {
            ccsSetInventory(r.inventory_item_id, target);
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

      logRows.push({
        timestamp: new Date(),
        mode: mode,
        sync_group: group,
        product_title: r.product_title,
        variant_title: r.variant_title,
        pack_size: pack,
        shopify_available: r.shopify_available,
        target_available: target,
        status: status,
        message: message
      });
    });
  });

  ccsWriteRows(CCS.tabs.mapping, CCS.headers.mapping, rows);
  ccsAppend(CCS.tabs.log, CCS.headers.log, logRows);
  SpreadsheetApp.getUi().alert(mode === 'live' ? 'Live pack sync complete. Check Sync Log.' : 'Preview complete. Check target_available and Sync Log.');
}

function ccsSetInventory(inventoryItemId, quantity) {
  const mutation = 'mutation InventorySet($input: InventorySetQuantitiesInput!) { inventorySetQuantities(input: $input) { userErrors { field message code } } }';
  const input = {
    name: 'available',
    reason: 'correction',
    referenceDocumentUri: 'can-count-sync://' + new Date().getTime(),
    quantities: [{ inventoryItemId: inventoryItemId, locationId: ccsGet('SHOPIFY_LOCATION_ID'), quantity: Number(quantity) }]
  };
  const data = ccsGraphql(mutation, { input: input });
  const errors = data.inventorySetQuantities.userErrors || [];
  if (errors.length) throw new Error(JSON.stringify(errors));
}

function ccsProcessRecentOrders() {
  let rows = ccsRefreshMappingRows(false);
  const activeByVariant = {};
  const groups = {};
  rows.forEach(r => {
    if (!ccsYes(r.active)) return;
    activeByVariant[String(r.variant_gid)] = r;
    const group = String(r.sync_group || '').trim();
    if (!groups[group]) groups[group] = [];
    groups[group].push(r);
  });

  const processed = {};
  ccsRead(CCS.tabs.orders).forEach(r => processed[String(r.shopify_order_id)] = true);

  const hours = Number(ccsGet('ORDER_LOOKBACK_HOURS', '24'));
  const since = new Date(Date.now() - hours * 60 * 60 * 1000).toISOString();
  const query = 'query Orders($q: String!) { orders(first: 50, query: $q, sortKey: CREATED_AT, reverse: true) { nodes { id name createdAt lineItems(first: 100) { nodes { quantity variant { id title } } } } } }';
  const data = ccsGraphql(query, { q: 'created_at:>=' + since });

  const adjustments = {};
  const orderRows = [];
  data.orders.nodes.forEach(order => {
    if (processed[String(order.id)]) return;
    order.lineItems.nodes.forEach(line => {
      const variantId = line.variant ? String(line.variant.id) : '';
      const sold = activeByVariant[variantId];
      if (!sold) return;
      const group = String(sold.sync_group || '').trim();
      const master = ccsFindMaster(groups[group] || []);
      if (!master) return;
      if (String(master.variant_gid) === String(sold.variant_gid)) return;

      const soldPack = ccsNumber(sold.pack_size);
      const masterPack = ccsNumber(master.pack_size);
      if (!soldPack || !masterPack) return;
      const unitsToDeductFromMaster = Math.ceil(Number(line.quantity || 0) * soldPack / masterPack);
      if (!adjustments[master.variant_gid]) adjustments[master.variant_gid] = { row: master, quantity: 0, notes: [] };
      adjustments[master.variant_gid].quantity += unitsToDeductFromMaster;
      adjustments[master.variant_gid].notes.push(order.name + ' sold ' + line.quantity + ' x ' + sold.variant_title);
    });
    orderRows.push({ timestamp: new Date(), shopify_order_id: order.id, shopify_order_name: order.name, status: 'processed', note: 'Checked by Can Count Sync' });
  });

  const logRows = [];
  Object.keys(adjustments).forEach(key => {
    const adj = adjustments[key];
    const current = ccsNumber(adj.row.shopify_available);
    const next = Math.max(0, current - adj.quantity);
    try {
      ccsSetInventory(adj.row.inventory_item_id, next);
      logRows.push({ timestamp: new Date(), mode: 'order_adjust', sync_group: adj.row.sync_group, product_title: adj.row.product_title, variant_title: adj.row.variant_title, pack_size: adj.row.pack_size, shopify_available: current, target_available: next, status: 'success', message: 'Master adjusted for pack order: ' + adj.notes.join('; ') });
    } catch (err) {
      logRows.push({ timestamp: new Date(), mode: 'order_adjust', sync_group: adj.row.sync_group, product_title: adj.row.product_title, variant_title: adj.row.variant_title, pack_size: adj.row.pack_size, shopify_available: current, target_available: next, status: 'error', message: err.message });
    }
  });

  ccsAppend(CCS.tabs.orders, CCS.headers.orders, orderRows);
  ccsAppend(CCS.tabs.log, CCS.headers.log, logRows);
  ccsRunPackSync('live');
}

function ccsInstallOrderPolling() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'ccsProcessRecentOrders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('ccsProcessRecentOrders').timeBased().everyMinutes(10).create();
  SpreadsheetApp.getUi().alert('10 minute order polling is installed. It will process recent orders and then sync related pack variants.');
}
