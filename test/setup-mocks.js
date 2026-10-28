jest.mock('@adobe/aio-sdk', () => ({
  Core: { Logger: jest.fn(() => ({ info: jest.fn(), debug: jest.fn(), error: jest.fn() })) }
}))
jest.mock('@adobe/aio-lib-core-auth', () => ({
  generateAccessToken: jest.fn(async () => ({ access_token: 'tok' }))
}))
// Skip real polling delays.
global.setTimeout = (fn) => { fn(); return 0 }
