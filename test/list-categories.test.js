require('./setup-mocks')
const { jsonResponse, routeFetch, commerceParams } = require('./helpers')
const { main } = require('../actions/list-categories/index.js')

describe('list-categories', () => {
  afterEach(() => jest.clearAllMocks())

  it('returns 200 with labelled categories from the derived API base', async () => {
    global.fetch = routeFetch([{
      match: 'https://na1-sandbox.api.commerce.adobe.com/TENANT/V1/categories/list',
      respond: jsonResponse({ items: [
        { id: 3, name: 'Men', path: '1/2/3', is_active: true },
        { id: 4, name: 'Jackets', path: '1/2/3/4', is_active: true }
      ] })
    }])
    const res = await main({ ...commerceParams })
    expect(res.statusCode).toBe(200)
    expect(res.body.categories).toEqual([
      { id: '3', name: 'Men', label: 'Men' },
      { id: '4', name: 'Jackets', label: 'Men > Jackets' }
    ])
  })

  it('returns 500 when Commerce is not configured', async () => {
    const res = await main({ COMMERCE_CLIENT_ID: 'cid', COMMERCE_CLIENT_SECRET: 's' })
    expect(res.statusCode).toBe(500)
    expect(res.body.error).toContain('COMMERCE_SANDBOX_ADMIN_URL')
  })

  it('returns 500 on a non-JSON (admin login page) response', async () => {
    global.fetch = routeFetch([{ match: '/V1/', respond: { ok: true, status: 200, text: async () => '<html>' } }])
    const res = await main({ ...commerceParams })
    expect(res.statusCode).toBe(500)
    expect(res.body.error).toContain('non-JSON')
  })
})
