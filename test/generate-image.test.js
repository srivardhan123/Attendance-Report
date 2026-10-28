require('./setup-mocks')
const { jsonResponse, routeFetch, psParams } = require('./helpers')
const { main } = require('../actions/generate-image/index.js')

const base = { ...psParams, productName: 'Linen Shirt', category: 'Shirts', referenceImageUrl: 'https://store.example.com/b.jpg' }

describe('generate-image', () => {
  afterEach(() => jest.clearAllMocks())

  it('returns 400 on missing params', async () => {
    const res = await main({ ...psParams })
    expect(res.statusCode).toBe(400)
    expect(res.body.error).toContain('Missing')
  })

  it('returns 200 with a Firefly image styled after the uploaded reference', async () => {
    global.fetch = routeFetch([
      { match: 'store.example.com', respond: jsonResponse({}, 200, { 'content-type': 'image/jpeg' }) },
      { match: '/v2/storage/image', respond: jsonResponse({ images: [{ id: 'up-1' }] }) },
      { match: '/v3/images/generate-async', respond: jsonResponse({ statusUrl: 'https://firefly-api.adobe.io/v3/status/j1' }) },
      { match: '/v3/status/j1', respond: jsonResponse({ status: 'succeeded', result: { outputs: [{ seed: 7, image: { url: 'https://s3.amazonaws.com/out.png' } }] } }) }
    ])
    const res = await main({ ...base, seed: 7 })
    expect(res.statusCode).toBe(200)
    expect(res.body).toMatchObject({ imageUrl: 'https://s3.amazonaws.com/out.png', seed: 7 })
    const genCall = global.fetch.mock.calls.find(([u]) => u.includes('generate-async'))
    const body = JSON.parse(genCall[1].body)
    expect(body.style.imageReference.source.uploadId).toBe('up-1')
    expect(body.prompt).toContain('Linen Shirt')
  })

  it('returns 500 when the Firefly job fails', async () => {
    global.fetch = routeFetch([
      { match: 'store.example.com', respond: jsonResponse({}, 200, { 'content-type': 'image/png' }) },
      { match: '/v2/storage/image', respond: jsonResponse({ images: [{ id: 'up-1' }] }) },
      { match: '/v3/images/generate-async', respond: jsonResponse({ statusUrl: 'https://firefly-api.adobe.io/v3/status/j1' }) },
      { match: '/v3/status/j1', respond: jsonResponse({ status: 'failed' }) }
    ])
    const res = await main(base)
    expect(res.statusCode).toBe(500)
    expect(res.body.error).toContain('failed')
  })
})
