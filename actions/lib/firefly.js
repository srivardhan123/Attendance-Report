const { generateAccessToken } = require('@adobe/aio-lib-core-auth')
const FIREFLY_SCOPES = ['openid', 'AdobeID', 'read_organizations', 'additional_info.projectedProductContext', 'firefly_api', 'firefly_enterprise']
const FF = 'https://firefly-api.adobe.io'

async function ffToken (params) {
  const token = await generateAccessToken({
    clientId: params.PHOTOSHOP_CLIENT_ID, // same S2S credential as Photoshop
    clientSecret: params.PHOTOSHOP_CLIENT_SECRET,
    orgId: params.PHOTOSHOP_ORG_ID,
    scopes: FIREFLY_SCOPES
  })
  return typeof token === 'string' ? token : token.access_token
}

// Upload raw image bytes -> uploadId. Reference it everywhere as { source: { uploadId } }.
async function ffUpload (token, clientId, bytes, contentType = 'image/png') {
  const r = await fetch(`${FF}/v2/storage/image`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'x-api-key': clientId, 'Content-Type': contentType },
    body: bytes
  })
  const d = await r.json()
  if (!r.ok) throw new Error(d.message || `Firefly upload ${r.status}`)
  return d.images[0].id
}

// POST an async job, then poll its statusUrl until terminal. `extraHeaders` carries x-model-version.
async function ffPoll (token, clientId, path, body, extraHeaders = {}, { tries = 90, delayMs = 2000 } = {}) {
  const headers = { Authorization: `Bearer ${token}`, 'x-api-key': clientId, 'Content-Type': 'application/json', ...extraHeaders }
  const r = await fetch(`${FF}${path}`, { method: 'POST', headers, body: JSON.stringify(body) })
  const submit = await r.json()
  if (!r.ok) throw new Error(submit.message || `Firefly ${r.status}: ${JSON.stringify(submit)}`)
  const statusUrl = submit.statusUrl
  if (!statusUrl) throw new Error('Firefly: no statusUrl in response ' + JSON.stringify(submit))
  for (let i = 0; i < tries; i++) {
    await new Promise((res) => setTimeout(res, delayMs))
    const s = await (await fetch(statusUrl, { headers })).json()
    if (s.status === 'succeeded') return s
    if (['failed', 'cancelled', 'timeout'].includes(s.status)) throw new Error('Firefly job ' + s.status + ': ' + JSON.stringify(s))
  }
  throw new Error('Firefly job timed out')
}

module.exports = { FIREFLY_SCOPES, ffToken, ffUpload, ffPoll }
