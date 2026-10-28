require('./setup-mocks')
const { jsonResponse, routeFetch, psParams } = require('./helpers')
const { main } = require('../actions/enhance-image/index.js')

describe('enhance-image', () => {
  afterEach(() => jest.clearAllMocks())

  it('returns 400 on missing imageUrl', async () => {
    const res = await main({ ...psParams })
    expect(res.statusCode).toBe(400)
    expect(res.body.error).toContain('Missing')
  })

  it('returns 200 after auto-tone then background removal', async () => {
    global.fetch = routeFetch([
      { match: 'photoshop-api.adobe.io/v2/edit', respond: jsonResponse({ _links: { self: { href: 'https://photoshop-api.adobe.io/v2/status/e1' } } }) },
      { match: '/v2/status/e1', respond: jsonResponse({ status: 'succeeded', outputs: [{ status: 'succeeded', destination: { url: 'https://photoshop-api.adobe.io/v2/short-url/x' } }] }) },
      { match: 'image.adobe.io/v2/remove-background', respond: jsonResponse({ statusUrl: 'https://firefly-tenant.adobe.io/v3/status/r1' }) },
      { match: 'firefly-tenant.adobe.io/v3/status/r1', respond: jsonResponse({ status: 'succeeded', result: { output: { destination: { url: 'https://s3.amazonaws.com/cutout.png' } } } }) }
    ])
    const res = await main({ ...psParams, imageUrl: 'https://s3.amazonaws.com/out.png' })
    expect(res.statusCode).toBe(200)
    expect(res.body).toEqual({ correctedUrl: 'https://photoshop-api.adobe.io/v2/short-url/x', finalUrl: 'https://s3.amazonaws.com/cutout.png' })
  })

  it('returns 500 when Photoshop rejects the job', async () => {
    global.fetch = routeFetch([{ match: '/v2/edit', respond: jsonResponse({ message: 'Bad image' }, 400) }])
    const res = await main({ ...psParams, imageUrl: 'https://s3.amazonaws.com/out.png' })
    expect(res.statusCode).toBe(500)
    expect(res.body.error).toBe('Bad image')
  })
})
