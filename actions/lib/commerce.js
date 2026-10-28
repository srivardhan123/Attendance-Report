const { generateAccessToken } = require('@adobe/aio-lib-core-auth')

const COMMERCE_SCOPES = [
  'openid', 'AdobeID', 'email', 'profile', 'additional_info.roles',
  'additional_info.projectedProductContext', 'commerce.accs'
]

// Admin dashboard URL -> REST API base. Swap admin.->api., keep ONLY the tenant segment.
function deriveApiBase (adminUrl) {
  if (!adminUrl) throw new Error('COMMERCE_SANDBOX_ADMIN_URL is not configured')
  const u = new URL(adminUrl)
  const host = u.host.replace(/^([^.]+)\.admin\./, '$1.api.')
  const tenant = u.pathname.split('/').filter(Boolean)[0]
  if (!tenant) throw new Error('Could not extract tenant id from COMMERCE_SANDBOX_ADMIN_URL')
  return `https://${host}/${tenant}`
}

// One REST call. `path` starts at /V1 (e.g. '/categories', '/products?searchCriteria...').
// store -> the `Store` header ('all' | 'default' | a store-view code). NO /rest/<store>/ segment.
async function commerceFetch (params, token, method, path, body, store = 'all') {
  const url = `${deriveApiBase(params.COMMERCE_SANDBOX_ADMIN_URL)}/V1${path}`
  const res = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      'x-api-key': params.__commerce_client_id,
      'x-gw-ims-org-id': params.__commerce_org_id,
      'Content-Type': 'application/json',
      Store: store
    },
    body: body ? JSON.stringify(body) : undefined
  })
  const text = await res.text()
  let data
  try { data = text ? JSON.parse(text) : {} } catch (_) {
    // Non-JSON body = wrong base URL (hit the admin page) — fail loudly, don't return empty.
    throw new Error(`Commerce ${method} ${path}: non-JSON response (${res.status}) — check deriveApiBase/base URL`)
  }
  if (!res.ok) {
    const err = new Error(data.message || `Commerce ${method} ${path} failed (${res.status})`)
    err.statusCode = res.status
    throw err
  }
  return data
}

// Gets a Commerce token. Prefers a dedicated COMMERCE_* credential; falls back to the
// project's shared S2S credential injected by `include-ims-credentials`.
// Also wires client id + org id onto params for commerceFetch.
async function commerceToken (params) {
  let token
  if (params.COMMERCE_CLIENT_ID && params.COMMERCE_CLIENT_SECRET) {
    token = await generateAccessToken({
      clientId: params.COMMERCE_CLIENT_ID,
      clientSecret: params.COMMERCE_CLIENT_SECRET,
      orgId: params.COMMERCE_ORG_ID,
      scopes: COMMERCE_SCOPES
    })
    params.__commerce_client_id = params.COMMERCE_CLIENT_ID
    params.__commerce_org_id = params.COMMERCE_ORG_ID
  } else if (params.__ims_oauth_s2s) {
    token = await generateAccessToken(params)
    params.__commerce_client_id = params.__ims_oauth_s2s.client_id
    params.__commerce_org_id = params.__ims_oauth_s2s.org_id
  } else {
    throw new Error('No Commerce credential configured (set COMMERCE_CLIENT_ID/SECRET/ORG_ID in .env)')
  }
  return typeof token === 'string' ? token : token.access_token
}

// Builds a public image URL for a gallery file path like "/m/b/mb01.jpg".
async function mediaBaseUrl (params, token) {
  const configs = await commerceFetch(params, token, 'GET', '/store/storeConfigs', null, 'default')
  const list = Array.isArray(configs) ? configs : [configs]
  const cfg = list.find(c => c.secure_base_media_url || c.base_media_url) || {}
  const base = cfg.secure_base_media_url || cfg.base_media_url || ''
  return base ? `${base.replace(/\/+$/, '')}/catalog/product` : ''
}

module.exports = { COMMERCE_SCOPES, deriveApiBase, commerceFetch, commerceToken, mediaBaseUrl }
