const { generateAccessToken } = require('@adobe/aio-lib-core-auth')
const PHOTOSHOP_SCOPES = ['openid', 'AdobeID', 'read_organizations', 'additional_info.projectedProductContext', 'creative_sdk']

async function psToken (params) {
  const token = await generateAccessToken({
    clientId: params.PHOTOSHOP_CLIENT_ID,
    clientSecret: params.PHOTOSHOP_CLIENT_SECRET,
    orgId: params.PHOTOSHOP_ORG_ID, // e.g. "<org id>@AdobeOrg"
    scopes: PHOTOSHOP_SCOPES
  })
  return typeof token === 'string' ? token : token.access_token
}

// POST a job, then poll the statusUrl the RESPONSE returns (never hardcode the status host).
async function psPoll (token, clientId, submitUrl, body, { tries = 60, delayMs = 2000 } = {}) {
  const headers = { Authorization: `Bearer ${token}`, 'x-api-key': clientId, 'Content-Type': 'application/json' }
  const r = await fetch(submitUrl, { method: 'POST', headers, body: JSON.stringify(body) })
  const submit = await r.json()
  if (!r.ok) throw new Error(submit.message || `Photoshop ${r.status}: ${JSON.stringify(submit.errors || submit)}`)
  const statusUrl = submit.statusUrl || (submit._links && submit._links.self && submit._links.self.href)
  if (!statusUrl) throw new Error('Photoshop: no statusUrl in response ' + JSON.stringify(submit))
  for (let i = 0; i < tries; i++) {
    await new Promise((res) => setTimeout(res, delayMs))
    const s = await (await fetch(statusUrl, { headers })).json()
    const status = s.status || (s.outputs && s.outputs[0] && s.outputs[0].status)
    if (status === 'succeeded') return s
    if (status === 'failed') throw new Error('Photoshop job failed: ' + JSON.stringify(s.errors || s))
  }
  throw new Error('Photoshop job timed out')
}

// Output URLs live in different places per endpoint — check them all.
function psOutputUrl (job) {
  const r = job.result || job
  const candidates = [
    r.outputs && r.outputs[0] && r.outputs[0].destination && r.outputs[0].destination.url,
    r.outputs && r.outputs[0] && r.outputs[0]._links && r.outputs[0]._links.renditions && r.outputs[0]._links.renditions[0] && r.outputs[0]._links.renditions[0].href,
    r.output && r.output.destination && r.output.destination.url,
    r.output && r.output.url,
    job.outputs && job.outputs[0] && job.outputs[0].destination && job.outputs[0].destination.url,
    job.output && job.output.destination && job.output.destination.url,
    job.output && job.output.url
  ]
  const url = candidates.find(Boolean)
  if (!url) throw new Error('Photoshop: no output URL in job result ' + JSON.stringify(job).slice(0, 500))
  return url
}

module.exports = { PHOTOSHOP_SCOPES, psToken, psPoll, psOutputUrl }
