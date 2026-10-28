// Shared mocks: route fetch calls by URL substring to canned responses.
function jsonResponse (body, status = 200, headers = {}) {
  return {
    ok: status >= 200 && status < 300,
    status,
    headers: { get: (h) => headers[h.toLowerCase()] || 'application/json' },
    json: async () => body,
    text: async () => JSON.stringify(body),
    arrayBuffer: async () => new TextEncoder().encode('img').buffer
  }
}

function routeFetch (routes) {
  return jest.fn(async (url, opts = {}) => {
    const method = (opts.method || 'GET').toUpperCase()
    const match = routes.find(r => String(url).includes(r.match) && (!r.method || r.method === method))
    if (!match) throw new Error(`Unmocked fetch ${method} ${url}`)
    return typeof match.respond === 'function' ? match.respond(url, opts) : match.respond
  })
}

const commerceParams = {
  COMMERCE_SANDBOX_ADMIN_URL: 'https://na1-sandbox.admin.commerce.adobe.com/TENANT/admin/admin/dashboard/',
  COMMERCE_CLIENT_ID: 'cid',
  COMMERCE_CLIENT_SECRET: 'secret',
  COMMERCE_ORG_ID: 'org@AdobeOrg'
}
const psParams = { PHOTOSHOP_CLIENT_ID: 'pcid', PHOTOSHOP_CLIENT_SECRET: 'ps', PHOTOSHOP_ORG_ID: 'org@AdobeOrg' }

module.exports = { jsonResponse, routeFetch, commerceParams, psParams }
