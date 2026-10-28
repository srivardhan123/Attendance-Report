const { Core } = require('@adobe/aio-sdk')
const { commerceFetch, commerceToken, mediaBaseUrl } = require('../lib/commerce')

const ORDER_PAGE_SIZE = 500
const MAX_ORDER_PAGES = 20

function searchCriteria (filters, extra = []) {
  const parts = []
  filters.forEach((f, i) => {
    parts.push(`searchCriteria[filter_groups][${i}][filters][0][field]=${encodeURIComponent(f.field)}`)
    parts.push(`searchCriteria[filter_groups][${i}][filters][0][value]=${encodeURIComponent(f.value)}`)
    parts.push(`searchCriteria[filter_groups][${i}][filters][0][condition_type]=${f.condition || 'eq'}`)
  })
  return parts.concat(extra).join('&')
}

function productImagePath (product) {
  const attr = (product.custom_attributes || []).find(a => a.attribute_code === 'image')
  if (attr && attr.value && attr.value !== 'no_selection') return attr.value
  const entries = (product.media_gallery_entries || []).filter(e => !e.disabled && e.media_type === 'image')
  const base = entries.find(e => (e.types || []).includes('image')) || entries[0]
  return base ? base.file : null
}

// Commerce has no bestsellers endpoint: sum qty_ordered over top-level order items.
async function aggregateSales (params, token, sinceIso) {
  const filters = sinceIso ? [{ field: 'created_at', value: sinceIso, condition: 'gteq' }] : []
  const sales = new Map() // product_id -> { qty, revenue }
  for (let page = 1; page <= MAX_ORDER_PAGES; page++) {
    const q = searchCriteria(filters, [`searchCriteria[pageSize]=${ORDER_PAGE_SIZE}`, `searchCriteria[currentPage]=${page}`])
    const data = await commerceFetch(params, token, 'GET', `/orders/items?${q}`)
    const items = data.items || []
    for (const item of items) {
      if (item.parent_item_id) continue // child rows of configurable/bundle products
      const key = String(item.product_id)
      const s = sales.get(key) || { qty: 0, revenue: 0 }
      s.qty += Number(item.qty_ordered || 0) - Number(item.qty_refunded || 0)
      s.revenue += Number(item.base_row_total || item.row_total || 0)
      sales.set(key, s)
    }
    if (items.length < ORDER_PAGE_SIZE || page * ORDER_PAGE_SIZE >= (data.total_count || 0)) break
  }
  return sales
}

async function main (params) {
  const logger = Core.Logger('get-bestsellers', { level: params.LOG_LEVEL || 'info' })
  try {
    logger.info('Action invoked')
    logger.debug('Params:', JSON.stringify({ categoryId: params.categoryId, days: params.days }))

    const missing = ['categoryId'].filter(p => !params[p])
    if (missing.length > 0) {
      return { statusCode: 400, body: { error: `Missing required params: ${missing.join(', ')}` } }
    }
    const days = Number(params.days) > 0 ? Number(params.days) : 90
    const limit = Math.min(Math.max(Number(params.limit) || 4, 1), 12)

    const token = await commerceToken(params)
    const productQuery = searchCriteria(
      [{ field: 'category_id', value: String(params.categoryId) }],
      ['searchCriteria[pageSize]=300']
    )
    const [productData, mediaBase] = await Promise.all([
      commerceFetch(params, token, 'GET', `/products?${productQuery}`),
      mediaBaseUrl(params, token)
    ])
    const products = (productData.items || []).filter(p => p.status !== 2) // 2 = disabled
    if (products.length === 0) {
      return { statusCode: 200, body: { bestsellers: [], basis: 'none', days } }
    }

    // Prefer recent sales; fall back to all-time, then to newest products if nothing sold.
    const since = new Date(Date.now() - days * 86400000).toISOString().slice(0, 19).replace('T', ' ')
    let basis = 'recent'
    let sales = await aggregateSales(params, token, since)
    const soldInCategory = s => products.some(p => (s.get(String(p.id)) || {}).qty > 0)
    if (!soldInCategory(sales)) {
      basis = 'all-time'
      sales = await aggregateSales(params, token, null)
      if (!soldInCategory(sales)) basis = 'newest'
    }

    const ranked = products
      .map(p => {
        const s = sales.get(String(p.id)) || { qty: 0, revenue: 0 }
        const file = productImagePath(p)
        return {
          id: p.id,
          sku: p.sku,
          name: p.name,
          price: p.price,
          unitsSold: s.qty,
          revenue: Math.round(s.revenue * 100) / 100,
          imageUrl: file && mediaBase ? `${mediaBase}${file.startsWith('/') ? '' : '/'}${file}` : null,
          createdAt: p.created_at
        }
      })
      .filter(p => p.imageUrl) // only products with a photo can serve as a style reference
      .sort((a, b) => basis === 'newest'
        ? String(b.createdAt).localeCompare(String(a.createdAt))
        : b.unitsSold - a.unitsSold || b.revenue - a.revenue)
      .filter(p => basis === 'newest' || p.unitsSold > 0)
      .slice(0, limit)

    logger.info(`Returning ${ranked.length} bestsellers (basis: ${basis})`)
    return { statusCode: 200, body: { bestsellers: ranked, basis, days } }
  } catch (error) {
    logger.error('Action failed:', error.message)
    return { statusCode: 500, body: { error: error.message } }
  }
}

exports.main = main
