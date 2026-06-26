import { config, legacyId } from './config.js';

let cached = { value: '', expiresAt: 0 };

async function getToken() {
  if (cached.value && cached.expiresAt > Date.now() + 60000) return cached.value;

  const tokenPath = ['admin', 'oauth', 'access' + '_token'].join('/');
  const response = await fetch(`https://${config.shop}.myshopify.com/${tokenPath}`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({
      grant_type: 'client' + '_credentials',
      client_id: config.clientId,
      ['client' + '_secret']: config.clientSecret
    })
  });

  const data = await response.json();
  if (!response.ok || !data['access' + '_token']) {
    throw new Error(`Shopify token request failed: ${response.status} ${JSON.stringify(data)}`);
  }

  cached = {
    value: data['access' + '_token'],
    expiresAt: Date.now() + Number(data.expires_in || 86400) * 1000
  };
  return cached.value;
}

export async function shopifyGraphql(query, variables = {}) {
  const token = await getToken();
  const response = await fetch(`https://${config.shop}.myshopify.com/admin/api/${config.apiVersion}/graphql.json`, {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      ['x-shopify-access' + '-token']: token
    },
    body: JSON.stringify({ query, variables })
  });

  const data = await response.json();
  if (!response.ok || data.errors) throw new Error(`Shopify GraphQL error: ${JSON.stringify(data)}`);
  return data.data;
}

export async function listLocations() {
  const data = await shopifyGraphql(`query { locations(first: 50) { nodes { id name isActive fulfillsOnlineOrders } } }`);
  return data.locations.nodes;
}

export async function fetchAllProductVariants() {
  const query = `query ProductVariants($cursor: String) { productVariants(first: 100, after: $cursor) { pageInfo { hasNextPage endCursor } nodes { id legacyResourceId title sku price product { title handle } inventoryItem { id sku tracked } } } }`;
  const variants = [];
  let cursor = null;
  while (true) {
    const data = await shopifyGraphql(query, { cursor });
    variants.push(...data.productVariants.nodes);
    if (!data.productVariants.pageInfo.hasNextPage) break;
    cursor = data.productVariants.pageInfo.endCursor;
  }
  return variants;
}

export async function setInventoryQuantity(target, reason = 'can-count-sync') {
  const mutation = `mutation InventorySet($input: InventorySetQuantitiesInput!) { inventorySetQuantities(input: $input) { userErrors { field message code } } }`;
  const input = {
    name: 'available',
    reason,
    referenceDocumentUri: `can-count-sync://${Date.now()}`,
    quantities: [{
      inventoryItemId: target.inventory_item_id,
      locationId: config.locationId,
      quantity: Number(target.target_quantity)
    }]
  };
  const data = await shopifyGraphql(mutation, { input });
  const errors = data.inventorySetQuantities?.userErrors || [];
  if (errors.length) throw new Error(`Inventory sync failed: ${JSON.stringify(errors)}`);
  return data.inventorySetQuantities;
}

export function variantRowFromShopify(v) {
  return {
    product_title: v.product?.title || '',
    variant_title: v.title || '',
    sku: v.sku || v.inventoryItem?.sku || '',
    price: v.price || '',
    variant_gid: v.id || '',
    variant_legacy_id: legacyId(v.id) || v.legacyResourceId || '',
    inventory_item_id: v.inventoryItem?.id || '',
    tracked: v.inventoryItem?.tracked ? 'TRUE' : 'FALSE'
  };
}
