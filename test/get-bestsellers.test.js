require('./setup-mocks')
const { jsonResponse, routeFetch, commerceParams } = require('./helpers')
const { main } = require('../actions/get-bestsellers/index.js')

const products = {
  items: [
    { id: 1, sku: 'A', name: 'Alpha', price: 20, status: 1, custom_attributes: [{ attribute_code: 'image', value: '/a/a.jpg' }] },
    { id: 2, sku: 'B', name: 'Beta', price: 30, status: 1, media_gallery_entries: [{ media_type: 'image', file: '/b/b.jpg', types: ['image'] }] },
    { id: 3, sku: 'C', name: 'NoPhoto', price: 10, status: 1 }
  ]
}
const storeConfigs = [{ secure_base_media_url: 'https://store.example.com/media/' }]

describe('get-bestsellers', () => {
  afterEach(() => jest.clearAllMocks())

  it('returns 400 when categoryId is missing', async () => {
    const res = await main({ ...commerceParams })
    expect(res.statusCode).toBe(400)
    expect(res.body.error).toContain('Missing')
  })

  it('returns 200 ranked by recent units sold, ignoring child rows', async () => {
    global.fetch = routeFetch([
      { match: '/V1/products?', respond: jsonResponse(products) },
      { match: '/V1/store/storeConfigs', respond: jsonResponse(storeConfigs) },
      { match: '/V1/orders/items', respond: jsonResponse({ total_count: 4, items: [
        { product_id: 1, qty_ordered: 2, row_total: 40 },
        { product_id: 2, qty_ordered: 5, row_total: 150 },
        { product_id: 2, qty_ordered: 5, parent_item_id: 9, row_total: 0 },
        { product_id: 99, qty_ordered: 50, row_total: 999 }
      ] }) }
    ])
    const res = await main({ ...commerceParams, categoryId: '4' })
    expect(res.statusCode).toBe(200)
    expect(res.body.basis).toBe('recent')
    expect(res.body.bestsellers.map(p => p.sku)).toEqual(['B', 'A'])
    expect(res.body.bestsellers[0]).toMatchObject({
      unitsSold: 5, revenue: 150, imageUrl: 'https://store.example.com/media/catalog/product/b/b.jpg'
    })
  })

  it('falls back to newest products when nothing has sold', async () => {
    global.fetch = routeFetch([
      { match: '/V1/products?', respond: jsonResponse(products) },
      { match: '/V1/store/storeConfigs', respond: jsonResponse(storeConfigs) },
      { match: '/V1/orders/items', respond: jsonResponse({ total_count: 0, items: [] }) }
    ])
    const res = await main({ ...commerceParams, categoryId: '4' })
    expect(res.statusCode).toBe(200)
    expect(res.body.basis).toBe('newest')
    expect(res.body.bestsellers).toHaveLength(2)
  })

  it('returns 500 when Commerce fails', async () => {
    global.fetch = routeFetch([{ match: '/V1/', respond: jsonResponse({ message: 'Unauthorized' }, 401) }])
    const res = await main({ ...commerceParams, categoryId: '4' })
    expect(res.statusCode).toBe(500)
    expect(res.body.error).toBe('Unauthorized')
  })
})
