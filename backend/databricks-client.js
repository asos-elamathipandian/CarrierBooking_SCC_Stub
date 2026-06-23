'use strict';

const { DBSQLClient } = require('@databricks/sql');
const { AzureCliCredential, InteractiveBrowserCredential, ChainedTokenCredential } = require('@azure/identity');

/**
 * Databricks SQL connector.
 * Reads from environment variables:
 *   DATABRICKS_HOST       — workspace hostname, e.g. adb-2908786112690092.12.azuredatabricks.net
 *   DATABRICKS_HTTP_PATH  — SQL Warehouse HTTP path, e.g. /sql/1.0/warehouses/abc1234
 *   DATABRICKS_TOKEN      — (optional) PAT — if not set, falls back to Azure AD auth
 *
 * Azure AD auth chain (used when PAT is absent or disabled):
 *   1. AzureCliCredential     — uses an existing `az login` session (no browser needed)
 *   2. InteractiveBrowserCredential — opens a browser login if az CLI isn't logged in
 */

// Databricks AAD resource ID (constant across all Azure Databricks workspaces)
const DATABRICKS_RESOURCE = '2ff814a6-3304-4ab8-85cb-cd0e6f879c1d';
const TENANT_ID           = '4af8322c-80ee-4819-a9ce-863d5afbea1c'; // ASOS Azure AD tenant

function getConfig() {
  const host     = process.env.DATABRICKS_HOST;
  const httpPath = process.env.DATABRICKS_HTTP_PATH;
  if (!host || !httpPath) return null;
  return { host, httpPath, token: process.env.DATABRICKS_TOKEN || null };
}

/**
 * Resolve the bearer token to use.
 * PAT takes priority; otherwise we obtain an AAD token via Azure CLI / browser.
 */
async function resolveToken(cfg) {
  if (cfg.token && !cfg.token.startsWith('REPLACE_')) return cfg.token;

  const credential = new ChainedTokenCredential(
    new AzureCliCredential({ tenantId: TENANT_ID }),
    new InteractiveBrowserCredential({ tenantId: TENANT_ID })
  );
  const tokenResponse = await credential.getToken(`${DATABRICKS_RESOURCE}/.default`);
  if (!tokenResponse?.token) throw new Error('Failed to obtain Azure AD token for Databricks');
  return tokenResponse.token;
}

/**
 * Run a SQL query and return rows as plain objects.
 */
async function query(sql, params) {
  const cfg = getConfig();
  if (!cfg) throw new Error('Databricks not configured — add DATABRICKS_HOST and DATABRICKS_HTTP_PATH to .env');

  const token  = await resolveToken(cfg);
  const client = new DBSQLClient();

  await client.connect({
    host:  cfg.host,
    path:  cfg.httpPath,
    token,
  });

  const session = await client.openSession({
    initialCatalog: 'supplychain',
    initialSchema:  'conformed',
  });

  try {
    const operation = await session.executeStatement(sql, {
      queryTimeout: 120,
      runAsync:     false,
      ...(params ? { namedParameters: params } : {}),
    });

    const result = await operation.fetchAll();
    await operation.close();
    return result;
  } finally {
    await session.close();
    await client.close();
  }
}

module.exports = { query, getConfig };
