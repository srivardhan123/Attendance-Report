require('./setup-mocks')
const { jsonResponse, routeFetch, commerceParams } = require('./helpers')
const { main } = require('../actions/publish-product/index.js')

const base = { ...commerceParams, name: 'Linen Shirt', price: 49, categoryId: '4', imageUrl: 'https://s3.amazonaws.com/cutout.png' }

describe('publish-product', () => {
  afterEach(() => jest.clearAllMocks())

  it('returns 400 on missing params', async () => {
    const res = await main({ ...commerceParams, name: 'X' })
    expect(res.statusCode).toBe(400)
    expect(res.body.error).toContain('Missing')
  })

  it('returns 400 on a non-positive price', async () => {
    const res = await main({ ...base, price: -1 })
    expect(res.statusCode).toBe(400)
    expect(res.body.error).toContain('price')
  })

  it('returns 200 and creates an in-stock product with the image attached', async () => {
    global.fetch = routeFetch([
      { match: 's3.amazonaws.com', respond: jsonResponse({}, 200, { 'content-type': 'image/png' }) },
      { match: '/V1/products/', method: 'POST', respond: jsonResponse(12) },
      { match: '/V1/products', method: 'POST', respond: jsonResponse({ id: 55, sku: 'x' }) }
    ])
    const res = await main(base)
    expect(res.statusCode).toBe(200)
    expect(res.body).toMatchObject({ id: 55, name: 'Linen Shirt', price: 49, qty: 100, imageAttached: true })
    expect(res.body.sku).toMatch(/^linen-shirt-/)
    const createCall = global.fetch.mock.calls.find(([u, o]) => u.endsWith('/V1/products') && o.method === 'POST')
    const { product } = JSON.parse(createCall[1].body)
    expect(product.extension_attributes.stock_item).toEqual({ qty: 100, is_in_stock: true, manage_stock: true })
    expect(product.extension_attributes.category_links[0].category_id).toBe('4')
    expect(product.status).toBe(1)
  })

  it('returns 500 when Commerce product creation fails', async () => {
    global.fetch = routeFetch([
      { match: 's3.amazonaws.com', respond: jsonResponse({}, 200, { 'content-type': 'image/png' }) },
      { match: '/V1/products', method: 'POST', respond: jsonResponse({ message: 'Server error' }, 500) }
    ])
    const res = await main(base)
    expect(res.statusCode).toBe(500)
    expect(res.body.error).toBe('Server error')
  })
})
