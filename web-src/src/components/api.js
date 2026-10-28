// Calls a web action with the shell's IMS token and org. Returns parsed JSON or throws.
export async function invokeAction (url, ims, body) {
  if (!url) return null
  const headers = { 'Content-Type': 'application/json' }
  if (ims && ims.token) headers.Authorization = `Bearer ${ims.token}`
  if (ims && ims.org) headers['x-gw-ims-org-id'] = ims.org
  const res = await fetch(url, { method: 'POST', headers, body: JSON.stringify(body || {}) })
  let data = null
  try { data = await res.json() } catch (e) { data = null }
  if (!res.ok) {
    const detail = data && (data.error || data.message)
    throw new Error(detail ? `${detail} (HTTP ${res.status})` : `Action failed: ${res.status}`)
  }
  return data
}

export const formatPrice = (value) =>
  new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(Number(value) || 0)
