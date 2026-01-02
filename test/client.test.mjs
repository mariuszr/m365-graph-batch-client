import { describe, expect, test, vi } from 'vitest'

import { getRetryAfterMs, M365GraphBatchClient, normalizeHeaders, toRelativeBatchUrl } from '..'
import { createPaginationHandler } from '../internal/pagination'
import { createRefreshTokenAccessTokenProvider } from '../internal/tokenProvider'

function createAxiosResponse({ status = 200, data, headers = {} }) {
  const normalizedHeaders = {}
  for (const [k, v] of Object.entries(headers || {})) normalizedHeaders[String(k).toLowerCase()] = String(v)

  return {
    status,
    data,
    headers: normalizedHeaders,
  }
}

function createMockSleep() {
  const calls = []
  const sleep = async (ms) => {
    calls.push(ms)
  }
  return { sleep, calls }
}

function createMockAxios(sequence) {
  const calls = []
  let idx = 0

  const axios = {
    async request(config) {
      calls.push(config)

      if (idx >= sequence.length) {
        throw new Error(`Mock axios out of responses (call ${calls.length})`)
      }

      const step = sequence[idx++]
      if (step.matcher) {
        expect(step.matcher(config)).toBe(true)
      }

      if (step.throw) throw step.throw
      return step.response
    },
  }

  return { axios, calls, remaining: () => sequence.length - idx }
}

describe('m365GraphBatchClient', () => {
  test('toRelativeBatchUrl strips origin and ensures leading slash', () => {
    expect(toRelativeBatchUrl('https://graph.microsoft.com/v1.0/users?$top=1')).toBe('/v1.0/users?$top=1')
    expect(toRelativeBatchUrl('/users')).toBe('/users')
    expect(toRelativeBatchUrl('users')).toBe('/users')
  })

  test('toRelativeBatchUrl keeps querystring when absolute url contains search', () => {
    expect(toRelativeBatchUrl('https://graph.microsoft.com/v1.0/users?x=1&y=2')).toBe('/v1.0/users?x=1&y=2')
  })

  test('getRetryAfterMs supports delta-seconds and HTTP-date', () => {
    expect(getRetryAfterMs({ 'Retry-After': '2' }, () => 0)).toBe(2000)

    const date = 'Thu, 01 Jan 1970 00:00:01 GMT'
    expect(getRetryAfterMs({ 'Retry-After': date }, () => 0)).toBe(1000)
  })

  test('getRetryAfterMs returns null for missing or invalid Retry-After', () => {
    expect(getRetryAfterMs({})).toBeNull()
    expect(getRetryAfterMs({ 'Retry-After': 'not-a-date-or-number' })).toBeNull()
  })

  test('getRetryAfterMs uses default now() and clamps to 0', () => {
    expect(getRetryAfterMs({ 'Retry-After': 'Thu, 01 Jan 1970 00:00:00 GMT' })).toBe(0)
  })

  test('normalizeHeaders filters null/undefined values', () => {
    expect(normalizeHeaders({ A: null, B: undefined, C: 'x' })).toEqual({ c: 'x' })
  })

  test('normalizeHeaders returns empty object for falsy input', () => {
    expect(normalizeHeaders(null)).toEqual({})
  })

  test('toRelativeBatchUrl throws for empty input', () => {
    expect(() => toRelativeBatchUrl('')).toThrow(/Request url is required/)
  })

  test('constructor throws when options is missing', () => {
    expect(() => new M365GraphBatchClient()).toThrow(/options is required/)
  })

  test('default sleep uses timers and resolves', async () => {
    vi.useFakeTimers()

    const promise = new Promise((resolve) => {
      setTimeout(resolve, 10)
      vi.advanceTimersByTime(10)
    })

    await promise
    vi.useRealTimers()
  })

  test('createDefaultSleep resolves after ms', async () => {
    vi.useFakeTimers()

    const { createDefaultSleep } = await import('../internal/utils')
    const sleep = createDefaultSleep()

    const promise = sleep(25)
    vi.advanceTimersByTime(25)
    await promise

    vi.useRealTimers()
  })

  test('default now() uses Date.now()', () => {
    const { axios } = createMockAxios([])
    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't' })

    const originalNow = Date.now
    Date.now = () => 123

    try {
      expect(client._now()).toBe(123)
    } finally {
      Date.now = originalNow
    }
  })

  test('_toFullUrl adds leading slash for relative paths', () => {
    const { axios } = createMockAxios([])
    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't' })

    expect(client._toFullUrl('users')).toBe('https://graph.microsoft.com/v1.0/users')
  })

  test('_toFullUrl prefixes to graphBaseUrl for /path', () => {
    const { axios } = createMockAxios([])
    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't' })

    expect(client._toFullUrl('/users')).toBe('https://graph.microsoft.com/v1.0/users')
  })

  test('_toFullUrl keeps absolute URL when origin matches graphBaseUrl', () => {
    const { axios } = createMockAxios([])
    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      graphBaseUrl: 'https://graph.microsoft.com/v1.0',
    })

    expect(client._toFullUrl('https://graph.microsoft.com/v1.0/users')).toBe('https://graph.microsoft.com/v1.0/users')
  })

  test('_toFullUrl does not throw if graphBaseUrl is invalid and url is absolute', () => {
    const { axios } = createMockAxios([])
    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't', graphBaseUrl: 'not-a-url' })

    expect(client._toFullUrl('https://evil.example/steal')).toBe('https://evil.example/steal')
  })

  test('constructor throws when getAccessToken/auth is missing', () => {
    expect(() => new M365GraphBatchClient({ axios: { request: async () => createAxiosResponse({}) } })).toThrow(
      /options\.getAccessToken is required/
    )
  })

  test('constructor throws when axiosInstance has no request', () => {
    expect(() => new M365GraphBatchClient({ axios: {}, getAccessToken: async () => 't' })).toThrow(
      /options\.axios\.request is required/
    )
  })

  test('constructor can use refresh_token auth when getAccessToken is not provided', async () => {
    const sleep = createMockSleep()

    const { axios, calls } = createMockAxios([
      {
        matcher: (config) => config.url.includes('/oauth2/v2.0/token') && config.method === 'POST',
        response: createAxiosResponse({ data: { access_token: 'rt-access', expires_in: 3600 } }),
      },
      {
        matcher: (config) => config.url === 'https://graph.microsoft.com/v1.0/$batch' && config.method === 'POST',
        response: createAxiosResponse({ data: { responses: [{ id: '1', status: 200, headers: {}, body: {} }] } }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      auth: {
        tenantId: 'tenant',
        clientId: 'client',
        clientSecret: 'secret',
        refreshToken: 'refresh',
      },
      sleep: sleep.sleep,
      initialBackoffMs: 0,
      jitterRatio: 0,
      maxBatchRetries: 0,
    })

    await client.batch([{ id: '1', url: '/x' }])

    expect(calls).toHaveLength(2)
    expect(calls[1].headers.authorization).toBe('Bearer rt-access')
  })

  test('constructor can lazy-require axios when not injected', () => {
    const client = new M365GraphBatchClient({ getAccessToken: async () => 't' })
    expect(client).toBeInstanceOf(M365GraphBatchClient)
  })

  test('constructor throws when axios dependency is missing and not injected', async () => {
    const axiosPath = require.resolve('axios')

    const saved = require.cache[axiosPath]
    delete require.cache[axiosPath]

    const originalLoad = require('node:module')._load
    require('node:module')._load = (request, parent, isMain) => {
      if (request === 'axios') {
        const err = new Error('not found')
        err.code = 'MODULE_NOT_FOUND'
        throw err
      }
      return originalLoad(request, parent, isMain)
    }

    try {
      const { M365GraphBatchClient: C } = require('../client.js')
      expect(() => new C({ getAccessToken: async () => 't' })).toThrow(/axios dependency not found/i)
    } finally {
      require('node:module')._load = originalLoad
      require.cache[axiosPath] = saved
    }
  })

  test('refresh_token provider caches token until expiry', async () => {
    const sleep = createMockSleep()

    const { axios, calls } = createMockAxios([
      {
        matcher: (config) => config.url.includes('/oauth2/v2.0/token') && config.method === 'POST',
        response: createAxiosResponse({ data: { access_token: 'rt-access', expires_in: 3600 } }),
      },
      {
        matcher: (config) => config.url === 'https://graph.microsoft.com/v1.0/$batch',
        response: createAxiosResponse({ data: { responses: [{ id: '1', status: 200, headers: {}, body: {} }] } }),
      },
      {
        matcher: (config) => config.url === 'https://graph.microsoft.com/v1.0/$batch',
        response: createAxiosResponse({ data: { responses: [{ id: '2', status: 200, headers: {}, body: {} }] } }),
      },
    ])

    // Provide deterministic now() so token is not considered expired.
    const now = 0
    const client = new M365GraphBatchClient({
      axios,
      auth: {
        tenantId: 'tenant',
        clientId: 'client',
        clientSecret: 'secret',
        refreshToken: 'refresh',
      },
      sleep: sleep.sleep,
      now: () => now,
      initialBackoffMs: 0,
      jitterRatio: 0,
      maxBatchRetries: 0,
    })

    await client.batch([{ id: '1', url: '/x' }])
    await client.batch([{ id: '2', url: '/x' }])

    // One token refresh + two batch calls.
    expect(calls).toHaveLength(3)
  })

  test('refresh_token provider returns cached token when not expired', async () => {
    const axios = {
      request: vi.fn(),
    }

    const now = vi.fn(() => 0)

    const getAccessToken = createRefreshTokenAccessTokenProvider({
      axios,
      tenantId: 'tenant',
      clientId: 'client',
      clientSecret: 'secret',
      refreshToken: 'refresh',
      now,
      clockSkewMs: 0,
    })

    axios.request.mockResolvedValueOnce(createAxiosResponse({ data: { access_token: 't1', expires_in: 3600 } }))

    expect(await getAccessToken()).toBe('t1')
    expect(axios.request).toHaveBeenCalledTimes(1)

    now.mockReturnValue(1000)
    expect(await getAccessToken()).toBe('t1')
    expect(axios.request).toHaveBeenCalledTimes(1)
  })

  test('refresh_token provider refreshes again after expiry', async () => {
    const sleep = createMockSleep()

    const { axios, calls } = createMockAxios([
      {
        matcher: (config) => config.url.includes('/oauth2/v2.0/token'),
        response: createAxiosResponse({ data: { access_token: 't1', expires_in: 1 } }),
      },
      {
        matcher: (config) => config.url === 'https://graph.microsoft.com/v1.0/$batch',
        response: createAxiosResponse({ data: { responses: [{ id: '1', status: 200, headers: {}, body: {} }] } }),
      },
      {
        matcher: (config) => config.url.includes('/oauth2/v2.0/token'),
        response: createAxiosResponse({ data: { access_token: 't2', expires_in: 3600 } }),
      },
      {
        matcher: (config) => config.url === 'https://graph.microsoft.com/v1.0/$batch',
        response: createAxiosResponse({ data: { responses: [{ id: '2', status: 200, headers: {}, body: {} }] } }),
      },
    ])

    let now = 0
    const client = new M365GraphBatchClient({
      axios,
      auth: {
        tenantId: 'tenant',
        clientId: 'client',
        clientSecret: 'secret',
        refreshToken: 'refresh',
        clockSkewMs: 0,
      },
      sleep: sleep.sleep,
      now: () => now,
      initialBackoffMs: 0,
      jitterRatio: 0,
      maxBatchRetries: 0,
    })

    await client.batch([{ id: '1', url: '/x' }])
    now = 2000
    await client.batch([{ id: '2', url: '/x' }])

    expect(calls[1].headers.authorization).toBe('Bearer t1')
    expect(calls[3].headers.authorization).toBe('Bearer t2')
    expect(calls).toHaveLength(4)
  })

  test('refresh_token provider throws on non-2xx token response', async () => {
    const sleep = createMockSleep()

    const { axios } = createMockAxios([
      {
        response: createAxiosResponse({ status: 400, data: { error: 'invalid_grant' } }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      auth: {
        tenantId: 'tenant',
        clientId: 'client',
        clientSecret: 'secret',
        refreshToken: 'refresh',
      },
      sleep: sleep.sleep,
      initialBackoffMs: 0,
      jitterRatio: 0,
      maxBatchRetries: 0,
    })

    await expect(client.batch([{ id: '1', url: '/x' }], { mode: 'strict' })).rejects.toThrow(
      /OAuth token refresh failed \(400\)/
    )
  })

  test('refresh_token provider throws when access_token is missing', async () => {
    const sleep = createMockSleep()

    const { axios } = createMockAxios([
      {
        response: createAxiosResponse({ data: { expires_in: 3600 } }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      auth: {
        tenantId: 'tenant',
        clientId: 'client',
        clientSecret: 'secret',
        refreshToken: 'refresh',
      },
      sleep: sleep.sleep,
      initialBackoffMs: 0,
      jitterRatio: 0,
      maxBatchRetries: 0,
    })

    await expect(client.batch([{ id: '1', url: '/x' }], { mode: 'strict' })).rejects.toThrow(/returned no access_token/)
  })

  test('refresh_token provider throws when expires_in is invalid', async () => {
    const sleep = createMockSleep()

    const { axios } = createMockAxios([
      {
        response: createAxiosResponse({ data: { access_token: 't', expires_in: 'nope' } }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      auth: {
        tenantId: 'tenant',
        clientId: 'client',
        clientSecret: 'secret',
        refreshToken: 'refresh',
      },
      sleep: sleep.sleep,
      initialBackoffMs: 0,
      jitterRatio: 0,
      maxBatchRetries: 0,
    })

    await expect(client.batch([{ id: '1', url: '/x' }], { mode: 'strict' })).rejects.toThrow(
      /returned invalid expires_in/
    )
  })

  test('createRefreshTokenAccessTokenProvider uses now() default', async () => {
    const { axios, calls } = createMockAxios([
      {
        response: createAxiosResponse({ data: { access_token: 't', expires_in: 3600 } }),
      },
    ])

    const getToken = createRefreshTokenAccessTokenProvider({
      axios,
      tenantId: 'tenant',
      clientId: 'client',
      clientSecret: 'secret',
      refreshToken: 'refresh',
    })

    await expect(getToken()).resolves.toBe('t')
    expect(calls).toHaveLength(1)
  })

  test('createRefreshTokenAccessTokenProvider always sets validateStatus', async () => {
    const { axios, calls } = createMockAxios([
      {
        matcher: (config) => typeof config.validateStatus === 'function',
        response: createAxiosResponse({ data: { access_token: 't', expires_in: 3600 } }),
      },
    ])

    const getToken = createRefreshTokenAccessTokenProvider({
      axios,
      tenantId: 'tenant',
      clientId: 'client',
      clientSecret: 'secret',
      refreshToken: 'refresh',
    })

    await getToken()
    expect(calls).toHaveLength(1)
    expect(calls[0].validateStatus(500)).toBe(true)
  })

  test('createRefreshTokenAccessTokenProvider validates required auth fields', () => {
    const { axios } = createMockAxios([])

    expect(() =>
      createRefreshTokenAccessTokenProvider({
        axios,
        tenantId: null,
        clientId: 'client',
        clientSecret: 'secret',
        refreshToken: 'refresh',
      })
    ).toThrow(/options\.auth\.tenantId is required/)

    expect(() =>
      createRefreshTokenAccessTokenProvider({
        axios,
        tenantId: 'tenant',
        clientId: null,
        clientSecret: 'secret',
        refreshToken: 'refresh',
      })
    ).toThrow(/options\.auth\.clientId is required/)

    expect(() =>
      createRefreshTokenAccessTokenProvider({
        axios,
        tenantId: 'tenant',
        clientId: 'client',
        clientSecret: null,
        refreshToken: 'refresh',
      })
    ).toThrow(/options\.auth\.clientSecret is required/)

    expect(() =>
      createRefreshTokenAccessTokenProvider({
        axios,
        tenantId: 'tenant',
        clientId: 'client',
        clientSecret: 'secret',
        refreshToken: null,
      })
    ).toThrow(/options\.auth\.refreshToken is required/)
  })

  test('createRefreshTokenAccessTokenProvider requires axios.request', () => {
    expect(() =>
      createRefreshTokenAccessTokenProvider({
        axios: null,
        tenantId: 'tenant',
        clientId: 'client',
        clientSecret: 'secret',
        refreshToken: 'refresh',
      })
    ).toThrow(/options\.axios\.request is required/)

    expect(() =>
      createRefreshTokenAccessTokenProvider({
        axios: {},
        tenantId: 'tenant',
        clientId: 'client',
        clientSecret: 'secret',
        refreshToken: 'refresh',
      })
    ).toThrow(/options\.axios\.request is required/)
  })

  test('createRefreshTokenAccessTokenProvider throws when token endpoint returns no response', async () => {
    const axios = {
      request: async () => null,
    }

    const getToken = createRefreshTokenAccessTokenProvider({
      axios,
      tenantId: 'tenant',
      clientId: 'client',
      clientSecret: 'secret',
      refreshToken: 'refresh',
    })

    await expect(getToken()).rejects.toThrow(/OAuth token refresh failed \(unknown\)/)
  })

  test('createRefreshTokenAccessTokenProvider uses plain string body in error', async () => {
    const axios = {
      request: async () => ({ status: 400, data: 'nope' }),
    }

    const getToken = createRefreshTokenAccessTokenProvider({
      axios,
      tenantId: 'tenant',
      clientId: 'client',
      clientSecret: 'secret',
      refreshToken: 'refresh',
    })

    await expect(getToken()).rejects.toThrow(/OAuth token refresh failed \(400\): nope/)
  })

  test('createRefreshTokenAccessTokenProvider uses JSON stringified body in error', async () => {
    const axios = {
      request: async () => ({ status: 400, data: { error: 'invalid_grant' } }),
    }

    const getToken = createRefreshTokenAccessTokenProvider({
      axios,
      tenantId: 'tenant',
      clientId: 'client',
      clientSecret: 'secret',
      refreshToken: 'refresh',
    })

    await expect(getToken()).rejects.toThrow(/OAuth token refresh failed \(400\): \{"error":"invalid_grant"\}/)
  })

  test('createRefreshTokenAccessTokenProvider uses deterministic now()', async () => {
    const axios = {
      request: async () => createAxiosResponse({ data: { access_token: 't', expires_in: 1 } }),
    }

    const now = () => 123

    const getToken = createRefreshTokenAccessTokenProvider({
      axios,
      tenantId: 'tenant',
      clientId: 'client',
      clientSecret: 'secret',
      refreshToken: 'refresh',
      now,
      clockSkewMs: 0,
    })

    await expect(getToken()).resolves.toBe('t')
  })

  test('createRefreshTokenAccessTokenProvider shares a single in-flight refresh', async () => {
    const { axios, calls } = createMockAxios([
      {
        response: createAxiosResponse({ data: { access_token: 't', expires_in: 3600 } }),
      },
    ])

    const getToken = createRefreshTokenAccessTokenProvider({
      axios,
      tenantId: 'tenant',
      clientId: 'client',
      clientSecret: 'secret',
      refreshToken: 'refresh',
    })

    const [a, b] = await Promise.all([getToken(), getToken()])
    expect(a).toBe('t')
    expect(b).toBe('t')
    expect(calls).toHaveLength(1)
  })

  test('createRefreshTokenAccessTokenProvider returns cached token when not expired', async () => {
    const { axios, calls } = createMockAxios([
      {
        response: createAxiosResponse({ data: { access_token: 't', expires_in: 3600 } }),
      },
    ])

    const getToken = createRefreshTokenAccessTokenProvider({
      axios,
      tenantId: 'tenant',
      clientId: 'client',
      clientSecret: 'secret',
      refreshToken: 'refresh',
      now: () => 0,
      clockSkewMs: 0,
    })

    expect(await getToken()).toBe('t')
    expect(await getToken()).toBe('t')
    expect(calls).toHaveLength(1)
  })

  test('constructor throws when axios dependency is missing', async () => {
    const Module = await import('node:module')

    const originalLoad = Module.default._load
    try {
      Module.default._load = function (request, parent, isMain) {
        if (request === 'axios') {
          const err = new Error('Cannot find module axios')
          err.code = 'MODULE_NOT_FOUND'
          throw err
        }
        return originalLoad.call(this, request, parent, isMain)
      }

      // dynamic import so it uses patched loader
      const { M365GraphBatchClient } = await import('..')

      expect(() => new M365GraphBatchClient({ getAccessToken: async () => 't' })).toThrow(/axios dependency not found/)
    } finally {
      Module.default._load = originalLoad
    }
  })

  test('falls back to relative nextLink when graphBaseUrl is invalid', async () => {
    const sleep = createMockSleep()

    const { axios } = createMockAxios([
      {
        response: createAxiosResponse({
          data: {
            responses: [
              {
                id: '1',
                status: 200,
                headers: {},
                body: {
                  value: [{ id: 1 }],
                  '@odata.nextLink': '/users?$skiptoken=abc',
                },
              },
            ],
          },
        }),
      },
      {
        // graphOrigin is null, so nextLink stays relative; _toFullUrl prefixes invalid base.
        matcher: (config) => config.url === 'not-a-url/users?$skiptoken=abc' && config.method === 'GET',
        response: createAxiosResponse({ data: { value: [{ id: 2 }] } }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      graphBaseUrl: 'not-a-url',
      sleep: sleep.sleep,
      jitterRatio: 0,
      maxBatchRetries: 0,
      initialBackoffMs: 0,
    })

    const out = await client.batch([{ id: '1', url: '/users' }])
    expect(out.responses['1'].body.value).toEqual([{ id: 1 }, { id: 2 }])
  })

  test('batch returns empty for empty input', async () => {
    const { axios } = createMockAxios([])
    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't' })

    await expect(client.batch([])).resolves.toEqual({ responses: {}, responseList: [] })

    // Cover request meta aggregation in _executeChunkWithRetries
    const { axios: axios2 } = createMockAxios([
      {
        response: createAxiosResponse({
          data: { responses: [{ id: '1', status: 200, headers: {}, body: { ok: true } }] },
        }),
      },
    ])

    const client2 = new M365GraphBatchClient({ axios: axios2, getAccessToken: async () => 't', maxBatchRetries: 0 })

    await expect(
      client2._executeChunkWithRetries([{ id: 1, url: '/x' }], { paginate: false, mode: 'partial' })
    ).resolves.toBeTruthy()
  })

  test('batch throws when requests is not an array', async () => {
    const { axios } = createMockAxios([])
    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't' })

    // eslint-disable-next-line no-undefined
    await expect(client.batch(undefined)).rejects.toThrow(/requests must be an array/)
  })

  test('_requestWithGlobalRetry retries on thrown errors then succeeds', async () => {
    const sleep = createMockSleep()

    const thrown = new Error('ECONNRESET')

    const { axios, calls } = createMockAxios([
      { throw: thrown },
      { throw: thrown },
      { response: createAxiosResponse({ data: { ok: true } }) },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      sleep: sleep.sleep,
      initialBackoffMs: 5,
      jitterRatio: 0,
      maxBatchRetries: 5,
    })

    await expect(client._requestWithGlobalRetry({ method: 'GET', url: '/x' })).resolves.toEqual({ ok: true })
    expect(calls).toHaveLength(3)
    expect(sleep.calls).toEqual([5, 10])

    // Cover non-retryable status helper
    expect(client._isRetryableStatus(400)).toBe(false)
  })

  test('_requestWithGlobalRetry throws RequestFailedError for non-retryable status', async () => {
    const { axios } = createMockAxios([{ response: createAxiosResponse({ status: 400, data: { error: 'nope' } }) }])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't', maxBatchRetries: 0 })

    await expect(client._requestWithGlobalRetry({ method: 'GET', url: '/x' })).rejects.toThrow(/Request failed \(400\)/)
  })

  test('_requestWithGlobalRetry blocks absolute URLs outside graph origin (SSRF protection)', async () => {
    const { axios, calls } = createMockAxios([{ response: createAxiosResponse({ status: 200, data: { ok: true } }) }])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      graphBaseUrl: 'https://graph.microsoft.com/v1.0',
      maxBatchRetries: 0,
    })

    await expect(client._requestWithGlobalRetry({ method: 'GET', url: 'https://evil.example/steal' })).rejects.toThrow(
      /origin mismatch/i
    )

    expect(calls).toHaveLength(0)
  })

  test('batch: partial mode marks external absolute subrequest as partial and does not call axios', async () => {
    const axios = {
      request: vi.fn(async () => {
        throw new Error('axios should not be called')
      }),
    }

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      graphBaseUrl: 'https://graph.microsoft.com/v1.0',
      maxBatchRetries: 0,
    })

    // cover _postBatchWithGlobalRetry origin validation (it should not run in this scenario)
    const postSpy = vi.spyOn(client, '_postBatchWithGlobalRetry')

    const out = await client.batch([{ id: '1', method: 'GET', url: 'https://evil.example/steal' }], {
      mode: 'partial',
    })

    expect(out.partial).toBe(true)
    expect(out.errors.some((e) => e.stage === 'batch' || e.stage === 'subrequest')).toBe(true)
    expect(out.responses['1']).toBeTruthy()
    expect(out.responses['1'].status).toBe(599)
    expect(axios.request).toHaveBeenCalledTimes(0)
    expect(postSpy).toHaveBeenCalledTimes(0)
  })

  test('_postBatchWithGlobalRetry annotates ORIGIN_MISMATCH with stage=subrequest', async () => {
    const axios = {
      request: vi.fn(async () => {
        throw new Error('axios should not be called')
      }),
    }

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      graphBaseUrl: 'https://graph.microsoft.com/v1.0',
      maxBatchRetries: 0,
    })

    await expect(
      client._postBatchWithGlobalRetry([{ id: '1', url: 'https://evil.example/steal' }])
    ).rejects.toMatchObject({
      code: 'ORIGIN_MISMATCH',
      stage: 'subrequest',
    })

    expect(axios.request).toHaveBeenCalledTimes(0)
  })

  test('batch: strict mode throws for external absolute subrequest', async () => {
    const axios = {
      request: vi.fn(async () => {
        throw new Error('axios should not be called')
      }),
    }

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      graphBaseUrl: 'https://graph.microsoft.com/v1.0',
      maxBatchRetries: 0,
    })

    await expect(
      client.batch([{ id: '1', method: 'GET', url: 'https://evil.example/steal' }], { mode: 'strict' })
    ).rejects.toThrow(/origin mismatch/i)

    expect(axios.request).toHaveBeenCalledTimes(0)
  })

  test('batch: partial mode filters off-origin subrequests and still runs $batch for valid ones', async () => {
    const axios = {
      request: vi.fn(async (_config) =>
        createAxiosResponse({
          data: {
            responses: [{ id: '2', status: 200, headers: {}, body: { ok: true } }],
          },
        })
      ),
    }

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      graphBaseUrl: 'https://graph.microsoft.com/v1.0',
      maxBatchRetries: 0,
    })

    const out = await client.batch(
      [
        { id: '1', method: 'GET', url: 'https://evil.example/steal' },
        { id: '2', method: 'GET', url: '/users?$top=1' },
      ],
      { mode: 'partial' }
    )

    expect(out.partial).toBe(true)
    expect(out.responses['1']?.status).toBe(599)
    expect(out.responses['2']?.status).toBe(200)

    // Should call axios once for the valid request.
    expect(axios.request).toHaveBeenCalledTimes(1)
    expect(axios.request.mock.calls[0][0].url).toBe('https://graph.microsoft.com/v1.0/$batch')

    // Ensure we only sent the valid subrequest in payload.
    expect(axios.request.mock.calls[0][0].data.requests).toHaveLength(1)
    expect(axios.request.mock.calls[0][0].data.requests[0].id).toBe('2')
  })

  test('_requestWithGlobalRetry retries on retryable status + Retry-After then succeeds', async () => {
    const sleep = createMockSleep()

    const { axios, calls } = createMockAxios([
      { response: createAxiosResponse({ status: 429, headers: { 'Retry-After': '2' }, data: 'slow down' }) },
      { response: createAxiosResponse({ status: 200, data: { ok: true } }) },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      sleep: sleep.sleep,
      initialBackoffMs: 1,
      jitterRatio: 0,
      maxBatchRetries: 5,
    })

    await expect(client._requestWithGlobalRetry({ method: 'GET', url: '/x' })).resolves.toEqual({ ok: true })
    expect(calls).toHaveLength(2)
    expect(sleep.calls).toEqual([2000])
  })

  test('_requestWithGlobalRetry throws RequestExceededRetriesError after max retries', async () => {
    const { axios } = createMockAxios([
      { response: createAxiosResponse({ status: 429, headers: { 'Retry-After': '0' }, data: 'slow down' }) },
      { response: createAxiosResponse({ status: 429, headers: { 'Retry-After': '0' }, data: 'slow down' }) },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      initialBackoffMs: 0,
      jitterRatio: 0,
      maxBatchRetries: 1,
    })

    await expect(client._requestWithGlobalRetry({ method: 'GET', url: '/x' })).rejects.toThrow(
      /Request exceeded retries \(last status 429\)/
    )
  })

  test('_postBatchWithGlobalRetry builds payload, normalizes headers, and validates shape', async () => {
    const { axios, calls } = createMockAxios([
      {
        matcher: (config) =>
          config.method === 'POST' &&
          config.url === 'https://graph.microsoft.com/v1.0/$batch' &&
          config.data?.requests?.[0]?.url === '/users' &&
          config.data?.requests?.[0]?.method === 'GET',
        response: createAxiosResponse({
          status: 200,
          data: {
            responses: [
              {
                id: '1',
                status: 200,
                headers: { 'Retry-After': 1 },
                body: { ok: true },
              },
            ],
          },
        }),
      },
    ])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't', maxBatchRetries: 0 })

    const out = await client._postBatchWithGlobalRetry([{ id: '1', method: 'GET', url: '/users' }])
    expect(calls).toHaveLength(1)
    expect(out.responses[0].headers).toEqual({ 'retry-after': '1' })

    // Invalid shape
    const badAxios = { request: async () => createAxiosResponse({ data: { nope: true } }) }
    const badClient = new M365GraphBatchClient({ axios: badAxios, getAccessToken: async () => 't' })

    await expect(badClient._postBatchWithGlobalRetry([{ id: '1', url: '/x' }])).rejects.toThrow(
      /Invalid \$batch response shape/
    )
  })

  test('executeChunkWithRetries retries only retryable subresponses, uses Retry-After, and paginates', async () => {
    const sleep = createMockSleep()

    const { axios, calls } = createMockAxios([
      // initial $batch
      {
        response: createAxiosResponse({
          data: {
            responses: [
              {
                id: '1',
                status: 429,
                headers: { 'Retry-After': '1' },
                body: { error: 'rate limit' },
              },
              {
                id: '2',
                status: 200,
                headers: {},
                body: { value: [{ id: 1 }], '@odata.nextLink': '/next' },
              },
            ],
          },
        }),
      },
      // retry $batch only for id=1
      {
        matcher: (config) => config.data.requests.length === 1 && config.data.requests[0].id === '1',
        response: createAxiosResponse({
          data: {
            responses: [
              {
                id: '1',
                status: 200,
                headers: {},
                body: { ok: true },
              },
            ],
          },
        }),
      },
      // pagination GET for id=2
      {
        matcher: (config) => config.method === 'GET' && config.url === 'https://graph.microsoft.com/v1.0/next',
        response: createAxiosResponse({ data: { value: [{ id: 2 }], '@odata.nextLink': '/next2' } }),
      },
      {
        matcher: (config) => config.method === 'GET' && config.url === 'https://graph.microsoft.com/v1.0/next2',
        response: createAxiosResponse({ data: { value: [] } }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      sleep: sleep.sleep,
      initialBackoffMs: 0,
      jitterRatio: 0,
      maxBatchRetries: 0,
      maxSubrequestRetries: 2,
    })

    const out = await client._executeChunkWithRetries(
      [
        { id: '1', method: 'GET', url: '/a' },
        { id: '2', method: 'GET', url: '/b' },
      ],
      { paginate: true, mode: 'partial' }
    )

    expect(calls).toHaveLength(3)
    expect(sleep.calls).toEqual([1000])
    expect(out.partial).toBe(true)
    expect(out.responseList[1].body.value).toEqual([{ id: 1 }])
  })

  test('batch sends a single $batch request and returns ordered responses', async () => {
    const { axios, calls } = createMockAxios([
      {
        matcher: (config) => config.url === 'https://graph.microsoft.com/v1.0/$batch' && config.method === 'POST',
        response: createAxiosResponse({
          data: {
            responses: [
              { id: '2', status: 200, headers: {}, body: { ok: true, id: 2 } },
              { id: '1', status: 200, headers: {}, body: { ok: true, id: 1 } },
            ],
          },
        }),
      },
    ])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't' })

    const out = await client.batch([
      { id: '1', method: 'GET', url: '/users' },
      { id: '2', method: 'GET', url: '/groups' },
    ])

    expect(calls).toHaveLength(1)
    expect(out.responseList).toHaveLength(2)
    expect(out.responseList[0].id).toBe('1')
    expect(out.responseList[1].id).toBe('2')
    expect(out.responses['1'].body.id).toBe(1)
    expect(out.responses['2'].body.id).toBe(2)
  })

  test('strict mode returns legacy shape (no partial/errors)', async () => {
    const { axios } = createMockAxios([
      {
        response: createAxiosResponse({ data: { responses: [{ id: '1', status: 200, headers: {}, body: {} }] } }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      maxBatchRetries: 0,
      initialBackoffMs: 0,
      jitterRatio: 0,
    })

    const out = await client.batch([{ id: '1', url: '/x' }], { mode: 'strict', paginate: false })

    expect(out.partial).toBeUndefined()
    expect(out.errors).toBeUndefined()
    expect(out.responses['1'].status).toBe(200)
  })

  test('default retryableStatuses includes 429, excludes 418', async () => {
    const { axios } = createMockAxios([])
    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't' })

    expect(client._isRetryableStatus(429)).toBe(true)
    expect(client._isRetryableStatus(418)).toBe(false)
  })

  test('custom retryableStatuses are respected', async () => {
    const { axios } = createMockAxios([])
    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't', retryableStatuses: [418] })

    expect(client._isRetryableStatus(418)).toBe(true)
    expect(client._isRetryableStatus(429)).toBe(false)
  })

  test('_computeBackoffMs uses backoff helper', () => {
    const { axios } = createMockAxios([])
    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      initialBackoffMs: 100,
      maxBackoffMs: 100,
      jitterRatio: 0,
    })

    expect(client._computeBackoffMs(1)).toBe(100)
  })

  test('_getWithGlobalRetry delegates to _requestWithGlobalRetry', async () => {
    const { axios } = createMockAxios([
      {
        response: createAxiosResponse({ status: 200, data: { ok: true } }),
      },
    ])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't', maxBatchRetries: 0 })
    await expect(client._getWithGlobalRetry('/me')).resolves.toEqual({ ok: true })
  })

  test('strict chunk execution returns legacy shape', async () => {
    const { axios } = createMockAxios([
      {
        response: createAxiosResponse({
          data: { responses: [{ id: '1', status: 200, headers: {}, body: { ok: true } }] },
        }),
      },
    ])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't', maxBatchRetries: 0 })

    const out = await client._executeChunkWithRetries([{ id: '1', url: '/x' }], { paginate: false, mode: 'strict' })

    expect(out).toEqual({
      responsesById: { 1: { id: '1', status: 200, headers: {}, body: { ok: true } } },
      responseList: [{ id: '1', status: 200, headers: {}, body: { ok: true } }],
    })
  })

  test('splits requests into chunks of maxRequestsPerBatch', async () => {
    const { axios, calls } = createMockAxios([
      {
        matcher: (config) => config.data.requests.length === 2,
        response: createAxiosResponse({
          data: {
            responses: [
              { id: '1', status: 200, headers: {}, body: {} },
              { id: '2', status: 200, headers: {}, body: {} },
            ],
          },
        }),
      },
      {
        matcher: (config) => config.data.requests.length === 1,
        response: createAxiosResponse({
          data: {
            responses: [{ id: '3', status: 200, headers: {}, body: {} }],
          },
        }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      maxRequestsPerBatch: 2,
    })

    const out = await client.batch([
      { id: '1', url: '/a' },
      { id: '2', url: '/b' },
      { id: '3', url: '/c' },
    ])

    expect(calls).toHaveLength(2)
    expect(out.responseList).toHaveLength(3)
  })

  test('validateStatus is always provided and returns true', async () => {
    const { axios, calls } = createMockAxios([
      {
        matcher: (config) => typeof config.validateStatus === 'function' && config.validateStatus(500) === true,
        response: createAxiosResponse({ data: { responses: [{ id: '1', status: 200, headers: {}, body: {} }] } }),
      },
    ])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't' })

    await client.batch([{ id: '1', url: '/x' }])
    expect(calls).toHaveLength(1)
  })

  test('_postBatchWithGlobalRetry throws when requestChunk exceeds maxRequestsPerBatch', async () => {
    const { axios } = createMockAxios([])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't', maxRequestsPerBatch: 1 })

    await expect(
      client._postBatchWithGlobalRetry([
        { id: '1', url: '/a' },
        { id: '2', url: '/b' },
      ])
    ).rejects.toThrow(/Batch request size exceeds 1/)
  })

  test('global retry: retries whole $batch on 429 with Retry-After', async () => {
    const sleep = createMockSleep()

    const { axios, calls } = createMockAxios([
      {
        response: createAxiosResponse({ status: 429, headers: { 'Retry-After': '3' }, data: 'throttle' }),
      },
      {
        response: createAxiosResponse({
          data: { responses: [{ id: '1', status: 200, headers: {}, body: { ok: true } }] },
        }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      sleep: sleep.sleep,
      maxBatchRetries: 2,
    })

    const out = await client.batch([{ id: '1', url: '/x' }])
    expect(calls).toHaveLength(2)
    expect(sleep.calls).toEqual([3000])
    expect(out.responses['1'].status).toBe(200)
  })

  test('requestWithGlobalRetry: throws detailed error for non-retryable status', async () => {
    const { axios } = createMockAxios([
      {
        response: createAxiosResponse({ status: 400, data: { error: { message: 'bad' } } }),
      },
    ])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't', maxBatchRetries: 0 })

    await expect(client._requestWithGlobalRetry({ method: 'GET', url: '/x' })).rejects.toThrow(/Request failed \(400\)/)
  })

  test('requestWithGlobalRetry: returns null for 204 with empty body', async () => {
    const { axios } = createMockAxios([
      {
        response: createAxiosResponse({ status: 204, data: undefined }),
      },
    ])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't' })

    await expect(client._requestWithGlobalRetry({ method: 'GET', url: '/x' })).resolves.toBeNull()
  })

  test('requestWithGlobalRetry: does not sleep when backoff is 0', async () => {
    const sleep = createMockSleep()

    const { axios } = createMockAxios([
      { throw: new Error('ECONNRESET') },
      {
        response: createAxiosResponse({ status: 204, data: undefined }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      sleep: sleep.sleep,
      jitterRatio: 0,
      maxBatchRetries: 2,
      initialBackoffMs: 0,
    })

    await expect(client._requestWithGlobalRetry({ method: 'GET', url: '/x' })).resolves.toBeNull()
    expect(sleep.calls).toEqual([])
  })

  test('requestWithGlobalRetry: uses plain string body in error message', async () => {
    const { axios } = createMockAxios([
      {
        response: createAxiosResponse({ status: 400, data: 'bad request' }),
      },
    ])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't', maxBatchRetries: 0 })

    await expect(client._requestWithGlobalRetry({ method: 'GET', url: '/x' })).rejects.toThrow(/bad request/)
  })

  test('requestWithGlobalRetry: stringifies empty body when data is undefined', async () => {
    const { axios } = createMockAxios([
      {
        response: createAxiosResponse({ status: 400, data: undefined }),
      },
    ])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't', maxBatchRetries: 0 })

    await expect(client._requestWithGlobalRetry({ method: 'GET', url: '/x' })).rejects.toThrow(/Request failed \(400\)/)
  })

  test('requestWithGlobalRetry: rethrows network error after exceeding maxBatchRetries', async () => {
    const { axios } = createMockAxios([
      {
        throw: new Error('ECONNRESET'),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      maxBatchRetries: 0,
      initialBackoffMs: 0,
    })

    await expect(client._requestWithGlobalRetry({ method: 'GET', url: '/x' })).rejects.toThrow(/ECONNRESET/)
  })

  test('global retry: retries on network error with exponential backoff', async () => {
    const sleep = createMockSleep()

    const { axios, calls } = createMockAxios([
      { throw: new Error('ECONNRESET') },
      { throw: new Error('ECONNRESET') },
      {
        response: createAxiosResponse({
          data: { responses: [{ id: '1', status: 200, headers: {}, body: {} }] },
        }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      sleep: sleep.sleep,
      jitterRatio: 0,
      maxBatchRetries: 5,
      initialBackoffMs: 10,
      maxBackoffMs: 1000,
    })

    await client.batch([{ id: '1', url: '/x' }])

    expect(calls).toHaveLength(3)
    expect(sleep.calls).toEqual([10, 20])
  })

  test('global retry: retries on 500/502 and uses Retry-After HTTP-date when present', async () => {
    const sleep = createMockSleep()

    const retryAfterDate = 'Thu, 01 Jan 1970 00:00:01 GMT'

    const { axios, calls } = createMockAxios([
      {
        response: createAxiosResponse({
          status: 500,
          headers: { 'Retry-After': retryAfterDate },
          data: { error: { message: 'server error' } },
        }),
      },
      {
        response: createAxiosResponse({
          status: 502,
          data: { error: { message: 'bad gateway' } },
        }),
      },
      {
        response: createAxiosResponse({
          data: { responses: [{ id: '1', status: 200, headers: {}, body: { ok: true } }] },
        }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      sleep: sleep.sleep,
      jitterRatio: 0,
      now: () => 0,
      maxBatchRetries: 5,
      initialBackoffMs: 10,
      maxBackoffMs: 1000,
    })

    const out = await client.batch([{ id: '1', url: '/x' }])

    expect(calls).toHaveLength(3)
    // 1st retry is driven by HTTP-date (1000ms), 2nd uses backoff (20ms).
    expect(sleep.calls).toEqual([1000, 20])
    expect(out.responses['1'].status).toBe(200)
  })

  test('global retry: throws after exceeding maxBatchRetries on retryable status', async () => {
    const { axios } = createMockAxios([
      { response: createAxiosResponse({ status: 500, data: 'e1' }) },
      { response: createAxiosResponse({ status: 500, data: 'e2' }) },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      jitterRatio: 0,
      maxBatchRetries: 1,
      initialBackoffMs: 0,
    })

    await expect(client.batch([{ id: '1', url: '/x' }], { mode: 'strict' })).rejects.toThrow(/exceeded retries/)
  })

  test('sub-request retry: retries only failed subrequests and respects Retry-After', async () => {
    const sleep = createMockSleep()

    const { axios, calls } = createMockAxios([
      {
        response: createAxiosResponse({
          data: {
            responses: [
              { id: '1', status: 200, headers: {}, body: { ok: 1 } },
              { id: '2', status: 429, headers: { 'Retry-After': '2' }, body: { error: 'throttle' } },
            ],
          },
        }),
      },
      {
        matcher: (config) => config.data.requests.length === 1 && config.data.requests[0].id === '2',
        response: createAxiosResponse({
          data: { responses: [{ id: '2', status: 200, headers: {}, body: { ok: 2 } }] },
        }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      sleep: sleep.sleep,
      jitterRatio: 0,
      maxBatchRetries: 0,
    })

    const out = await client.batch([
      { id: '1', url: '/a' },
      { id: '2', url: '/b' },
    ])

    expect(calls).toHaveLength(2)
    expect(sleep.calls).toEqual([2000])
    expect(out.responses['1'].body.ok).toBe(1)
    expect(out.responses['2'].body.ok).toBe(2)
  })

  test('sub-request retry: retries 500 and respects Retry-After HTTP-date', async () => {
    const sleep = createMockSleep()

    const retryAfterDate = 'Thu, 01 Jan 1970 00:00:01 GMT'

    const { axios, calls } = createMockAxios([
      {
        response: createAxiosResponse({
          data: {
            responses: [{ id: '1', status: 500, headers: { 'Retry-After': retryAfterDate }, body: { error: 'e' } }],
          },
        }),
      },
      {
        matcher: (config) => config.data.requests.length === 1 && config.data.requests[0].id === '1',
        response: createAxiosResponse({
          data: {
            responses: [{ id: '1', status: 200, headers: {}, body: { ok: true } }],
          },
        }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      sleep: sleep.sleep,
      now: () => 0,
      jitterRatio: 0,
      maxBatchRetries: 0,
    })

    const out = await client.batch([{ id: '1', url: '/users' }])

    expect(calls).toHaveLength(2)
    expect(sleep.calls).toEqual([1000])
    expect(out.responses['1'].status).toBe(200)
  })

  test('sub-request retry: uses exponential backoff when no Retry-After', async () => {
    const sleep = createMockSleep()

    const { axios } = createMockAxios([
      {
        response: createAxiosResponse({
          data: { responses: [{ id: '1', status: 503, headers: {}, body: { e: 1 } }] },
        }),
      },
      {
        response: createAxiosResponse({
          data: { responses: [{ id: '1', status: 503, headers: {}, body: { e: 2 } }] },
        }),
      },
      {
        response: createAxiosResponse({
          data: { responses: [{ id: '1', status: 200, headers: {}, body: { ok: true } }] },
        }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      sleep: sleep.sleep,
      jitterRatio: 0,
      initialBackoffMs: 10,
      maxBackoffMs: 1000,
      maxBatchRetries: 0,
    })

    const out = await client.batch([{ id: '1', url: '/x' }])
    expect(out.responses['1'].status).toBe(200)
    expect(sleep.calls).toEqual([10, 20])
  })

  test('sub-request retry: throws when exceeds maxSubrequestRetries', async () => {
    const sleep = createMockSleep()

    const { axios } = createMockAxios([
      {
        response: createAxiosResponse({
          data: { responses: [{ id: '1', status: 429, headers: { 'Retry-After': '0' }, body: {} }] },
        }),
      },
      {
        response: createAxiosResponse({
          data: { responses: [{ id: '1', status: 429, headers: { 'Retry-After': '0' }, body: {} }] },
        }),
      },
      {
        response: createAxiosResponse({
          data: { responses: [{ id: '1', status: 429, headers: { 'Retry-After': '0' }, body: {} }] },
        }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      sleep: sleep.sleep,
      maxSubrequestRetries: 1,
      maxBatchRetries: 0,
    })

    await expect(client.batch([{ id: '1', url: '/x' }], { mode: 'strict' })).rejects.toThrow(/exceeded retries/)
  })

  test('partial mode: does not throw when subrequest exceeds retries (returns last subresponse)', async () => {
    const sleep = createMockSleep()

    const { axios, calls } = createMockAxios([
      {
        response: createAxiosResponse({
          data: { responses: [{ id: '1', status: 429, headers: { 'Retry-After': '0' }, body: { e: 1 } }] },
        }),
      },
      {
        response: createAxiosResponse({
          data: { responses: [{ id: '1', status: 429, headers: { 'Retry-After': '0' }, body: { e: 2 } }] },
        }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      sleep: sleep.sleep,
      maxSubrequestRetries: 1,
      maxBatchRetries: 0,
      initialBackoffMs: 0,
      jitterRatio: 0,
    })

    const out = await client.batch([{ id: '1', url: '/x' }], { mode: 'partial' })

    expect(calls).toHaveLength(2)
    expect(out.partial).toBe(true)
    expect(out.errors).toHaveLength(1)
    expect(out.errors[0].id).toBe('1')
    expect(out.responses['1'].status).toBe(429)
    expect(out.responses['1'].body).toEqual({ e: 2 })
  })

  test('sub-request retry: uses max Retry-After when multiple subrequests are pending', async () => {
    const sleep = createMockSleep()

    const { axios } = createMockAxios([
      {
        response: createAxiosResponse({
          data: {
            responses: [
              { id: '1', status: 429, headers: { 'Retry-After': '1' }, body: {} },
              { id: '2', status: 429, headers: { 'Retry-After': '3' }, body: {} },
            ],
          },
        }),
      },
      {
        response: createAxiosResponse({
          data: {
            responses: [
              { id: '1', status: 200, headers: {}, body: { ok: 1 } },
              { id: '2', status: 200, headers: {}, body: { ok: 2 } },
            ],
          },
        }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      sleep: sleep.sleep,
      jitterRatio: 0,
      maxBatchRetries: 0,
    })

    const out = await client.batch([
      { id: '1', url: '/a' },
      { id: '2', url: '/b' },
    ])

    expect(out.responses['1'].status).toBe(200)
    expect(out.responses['2'].status).toBe(200)
    expect(sleep.calls).toEqual([3000])
  })

  test('sub-request retry: does not retry non-retryable subresponse status', async () => {
    const { axios, calls } = createMockAxios([
      {
        response: createAxiosResponse({
          data: {
            responses: [{ id: '1', status: 404, headers: {}, body: { error: 'not found' } }],
          },
        }),
      },
    ])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't', maxBatchRetries: 0 })

    const out = await client.batch([{ id: '1', url: '/x' }])

    expect(calls).toHaveLength(1)
    expect(out.responses['1'].status).toBe(404)
  })

  test('pagination: follows @odata.nextLink and aggregates value arrays', async () => {
    const { axios, calls } = createMockAxios([
      {
        response: createAxiosResponse({
          data: {
            responses: [
              {
                id: '1',
                status: 200,
                headers: {},
                body: {
                  value: [{ id: 1 }],
                  '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc',
                },
              },
            ],
          },
        }),
      },
      {
        matcher: (config) =>
          config.url === 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc' && config.method === 'GET',
        response: createAxiosResponse({
          data: {
            value: [{ id: 2 }],
            '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=def',
          },
        }),
      },
      {
        matcher: (config) =>
          config.url === 'https://graph.microsoft.com/v1.0/users?$skiptoken=def' && config.method === 'GET',
        response: createAxiosResponse({
          data: {
            value: [{ id: 3 }],
          },
        }),
      },
    ])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't' })

    const out = await client.batch([{ id: '1', url: '/users' }])
    expect(calls).toHaveLength(3)
    expect(out.responses['1'].body.value).toEqual([{ id: 1 }, { id: 2 }, { id: 3 }])
    expect(out.responses['1'].body['@odata.nextLink']).toBeUndefined()
  })

  test('pagination handler: default onError is used in partial mode', async () => {
    const handler = createPaginationHandler({
      getWithGlobalRetry: async () => 'not-json',
      graphOrigin: 'https://graph.microsoft.com',
      maxPaginationPages: 10,
    })

    const responseList = [
      {
        id: '1',
        status: 200,
        headers: {},
        body: {
          value: [{ id: 1 }],
          '@odata.nextLink': '/v1.0/users?$skiptoken=abc',
        },
      },
    ]

    const requestMetaById = {
      1: { method: 'GET' },
    }

    await expect(
      handler.paginateResponsesInPlace(responseList, requestMetaById, { mode: 'partial' })
    ).resolves.toBeUndefined()
    expect(responseList[0].body.value).toEqual([{ id: 1 }])
    expect(responseList[0].body['@odata.nextLink']).toBe('https://graph.microsoft.com/v1.0/users?$skiptoken=abc')
  })

  test('pagination: can be disabled via options.paginate=false', async () => {
    const { axios, calls } = createMockAxios([
      {
        response: createAxiosResponse({
          data: {
            responses: [
              {
                id: '1',
                status: 200,
                headers: {},
                body: {
                  value: [{ id: 1 }],
                  '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc',
                },
              },
            ],
          },
        }),
      },
    ])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't' })

    const out = await client.batch([{ id: '1', url: '/users' }], { paginate: false })
    expect(calls).toHaveLength(1)
    expect(out.responses['1'].body.value).toEqual([{ id: 1 }])
    expect(out.responses['1'].body['@odata.nextLink']).toBeDefined()
  })

  test('GET pagination uses global retry on throttling', async () => {
    const sleep = createMockSleep()

    const { axios, calls } = createMockAxios([
      {
        response: createAxiosResponse({
          data: {
            responses: [
              {
                id: '1',
                status: 200,
                headers: {},
                body: {
                  value: [{ id: 1 }],
                  '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc',
                },
              },
            ],
          },
        }),
      },
      {
        response: createAxiosResponse({ status: 429, headers: { 'Retry-After': '1' }, data: 'throttle' }),
      },
      {
        response: createAxiosResponse({ data: { value: [{ id: 2 }] } }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      sleep: sleep.sleep,
      maxBatchRetries: 2,
    })

    const out = await client.batch([{ id: '1', url: '/users' }])
    expect(calls).toHaveLength(3)
    expect(sleep.calls).toEqual([1000])
    expect(out.responses['1'].body.value).toEqual([{ id: 1 }, { id: 2 }])
  })

  test('does not auto-paginate non-GET requests', async () => {
    const { axios, calls } = createMockAxios([
      {
        response: createAxiosResponse({
          data: {
            responses: [
              {
                id: '1',
                status: 200,
                headers: {},
                body: {
                  value: [{ id: 1 }],
                  '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc',
                },
              },
            ],
          },
        }),
      },
    ])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't' })

    const out = await client.batch([{ id: '1', method: 'POST', url: '/users' }])
    expect(calls).toHaveLength(1)
    expect(out.responses['1'].body['@odata.nextLink']).toBeDefined()
  })

  test('pagination supports relative @odata.nextLink', async () => {
    const { axios, calls } = createMockAxios([
      {
        response: createAxiosResponse({
          data: {
            responses: [
              {
                id: '1',
                status: 200,
                headers: {},
                body: {
                  value: [{ id: 1 }],
                  '@odata.nextLink': '/users?$skiptoken=abc',
                },
              },
            ],
          },
        }),
      },
      {
        matcher: (config) =>
          config.url === 'https://graph.microsoft.com/users?$skiptoken=abc' && config.method === 'GET',
        response: createAxiosResponse({ data: { value: [{ id: 2 }] } }),
      },
    ])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't' })

    const out = await client.batch([{ id: '1', url: '/users' }])
    expect(calls).toHaveLength(2)
    expect(out.responses['1'].body.value).toEqual([{ id: 1 }, { id: 2 }])
  })

  test('pagination ignores responses not eligible for pagination', async () => {
    const { axios, calls } = createMockAxios([
      {
        response: createAxiosResponse({
          data: {
            responses: [
              // Non-2xx should be ignored by pagination; use non-retryable to avoid retry loops.
              {
                id: 'a',
                status: 404,
                headers: {},
                body: {
                  value: [{ id: 1 }],
                  '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc',
                },
              },
              // 2xx but body not object
              { id: 'b', status: 200, headers: {}, body: 'x' },
              // 2xx but value not an array
              {
                id: 'c',
                status: 200,
                headers: {},
                body: { value: 'x', '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc' },
              },
              // 2xx and value array but no nextLink
              { id: 'd', status: 200, headers: {}, body: { value: [{ id: 1 }] } },
            ],
          },
        }),
      },
    ])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't', maxBatchRetries: 0 })

    const out = await client.batch([
      { id: 'a', url: '/x' },
      { id: 'b', url: '/x' },
      { id: 'c', url: '/x' },
      { id: 'd', url: '/x' },
    ])

    expect(calls).toHaveLength(1)
    expect(out.responses['a'].status).toBe(404)
    expect(out.responses['b'].body).toBe('x')
    expect(out.responses['c'].body.value).toBe('x')
    expect(out.responses['d'].body.value).toEqual([{ id: 1 }])
  })

  test('pagination does not resolve absolute nextLink against graphOrigin', async () => {
    const { axios, calls } = createMockAxios([
      {
        response: createAxiosResponse({
          data: {
            responses: [
              {
                id: '1',
                status: 200,
                headers: {},
                body: {
                  value: [{ id: 1 }],
                  '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc',
                },
              },
            ],
          },
        }),
      },
      {
        matcher: (config) => config.url === 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc',
        response: createAxiosResponse({ data: { value: [{ id: 2 }] } }),
      },
    ])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't', maxBatchRetries: 0 })

    const out = await client.batch([{ id: '1', url: '/users' }])
    expect(calls).toHaveLength(2)
    expect(out.responses['1'].body.value).toEqual([{ id: 1 }, { id: 2 }])
  })

  test('paginateResponsesInPlace handles missing requestMetaById', async () => {
    const { axios } = createMockAxios([])
    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't' })

    await expect(
      client._paginateResponsesInPlace([
        {
          id: '1',
          status: 200,
          headers: {},
          body: { value: [{ id: 1 }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc' },
        },
      ])
    ).resolves.toBeUndefined()
  })

  test('pagination ignores page.value when it is not an array', async () => {
    const { axios, calls } = createMockAxios([
      {
        response: createAxiosResponse({
          data: {
            responses: [
              {
                id: '1',
                status: 200,
                headers: {},
                body: {
                  value: [{ id: 1 }],
                  '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc',
                },
              },
            ],
          },
        }),
      },
      {
        response: createAxiosResponse({
          data: {
            value: 'not-an-array',
          },
        }),
      },
    ])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't', maxBatchRetries: 0 })

    const out = await client.batch([{ id: '1', url: '/users' }])
    expect(calls).toHaveLength(2)
    expect(out.responses['1'].body.value).toEqual([{ id: 1 }])
  })

  test('throws on invalid $batch response shape', async () => {
    const { axios } = createMockAxios([
      {
        response: createAxiosResponse({ data: { notResponses: [] } }),
      },
    ])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't' })

    await expect(client.batch([{ id: '1', url: '/users' }], { mode: 'strict' })).rejects.toThrow(
      /Invalid \$batch response shape/
    )
  })

  test('throws on non-retryable HTTP status (e.g. 401)', async () => {
    const { axios } = createMockAxios([
      {
        response: createAxiosResponse({ status: 401, data: { error: { message: 'Unauthorized' } } }),
      },
    ])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't', maxBatchRetries: 0 })

    await expect(client.batch([{ id: '1', url: '/users' }], { mode: 'strict' })).rejects.toThrow(
      /Request failed \(401\)/
    )
  })

  test('duplicate request ids: last response wins in responses map', async () => {
    const { axios } = createMockAxios([
      {
        response: createAxiosResponse({
          data: {
            responses: [
              { id: '1', status: 200, headers: {}, body: { seq: 1 } },
              { id: '1', status: 200, headers: {}, body: { seq: 2 } },
            ],
          },
        }),
      },
    ])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't' })

    const out = await client.batch([{ id: '1', url: '/a' }])
    expect(out.responses['1'].body.seq).toBe(2)
  })

  test('missing subresponse is treated as retryable and retried', async () => {
    const { axios, calls } = createMockAxios([
      {
        response: createAxiosResponse({
          data: {
            responses: [],
          },
        }),
      },
      {
        response: createAxiosResponse({
          data: {
            responses: [{ id: '1', status: 200, headers: {}, body: { ok: true } }],
          },
        }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      maxBatchRetries: 0,
      initialBackoffMs: 0,
    })

    const out = await client.batch([{ id: '1', url: '/a' }])
    expect(calls).toHaveLength(2)
    expect(out.responses['1'].body.ok).toBe(true)
  })

  test('missing subresponse: throws after exceeding maxSubrequestRetries', async () => {
    const sleep = createMockSleep()

    const { axios, calls } = createMockAxios([
      { response: createAxiosResponse({ data: { responses: [] } }) },
      { response: createAxiosResponse({ data: { responses: [] } }) },
      { response: createAxiosResponse({ data: { responses: [] } }) },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      sleep: sleep.sleep,
      jitterRatio: 0,
      maxBatchRetries: 0,
      maxSubrequestRetries: 1,
      initialBackoffMs: 0,
    })

    await expect(client.batch([{ id: '1', url: '/a' }], { mode: 'strict' })).rejects.toThrow(/exceeded retries/)
    // initial + 1 retry; second retry attempt triggers the throw before calling $batch.
    expect(calls).toHaveLength(2)
  })

  test('partial mode: missing subresponse creates synthetic 599 response', async () => {
    const sleep = createMockSleep()

    const { axios, calls } = createMockAxios([
      { response: createAxiosResponse({ data: { responses: [] } }) },
      { response: createAxiosResponse({ data: { responses: [] } }) },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      sleep: sleep.sleep,
      jitterRatio: 0,
      maxBatchRetries: 0,
      maxSubrequestRetries: 1,
      initialBackoffMs: 0,
    })

    const out = await client.batch([{ id: '1', url: '/a' }], { mode: 'partial' })

    expect(calls).toHaveLength(2)
    expect(out.partial).toBe(true)
    expect(out.errors).toHaveLength(1)
    expect(out.responses['1'].status).toBe(599)
    expect(out.responses['1'].body.error.code).toBe('SubrequestExceededRetries')
  })

  test('pagination retries on network error during nextLink fetch', async () => {
    const sleep = createMockSleep()

    const { axios, calls } = createMockAxios([
      {
        // initial $batch
        response: createAxiosResponse({
          data: {
            responses: [
              {
                id: '1',
                status: 200,
                headers: {},
                body: {
                  value: [{ id: 1 }],
                  '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc',
                },
              },
            ],
          },
        }),
      },
      {
        // network error page 2
        throw: new Error('ECONNRESET'),
      },
      {
        // retry success
        response: createAxiosResponse({ data: { value: [{ id: 2 }] } }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      sleep: sleep.sleep,
      jitterRatio: 0,
      maxBatchRetries: 2,
      initialBackoffMs: 10,
      maxBackoffMs: 1000,
    })

    const out = await client.batch([{ id: '1', url: '/users' }])
    expect(calls).toHaveLength(3)
    expect(sleep.calls).toEqual([10])
    expect(out.responses['1'].body.value).toEqual([{ id: 1 }, { id: 2 }])
  })

  test('pagination retries on retryable HTTP status (e.g. 502) during nextLink fetch', async () => {
    const sleep = createMockSleep()

    const { axios, calls } = createMockAxios([
      {
        response: createAxiosResponse({
          data: {
            responses: [
              {
                id: '1',
                status: 200,
                headers: {},
                body: {
                  value: [{ id: 1 }],
                  '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc',
                },
              },
            ],
          },
        }),
      },
      {
        response: createAxiosResponse({
          status: 502,
          data: { error: { message: 'bad gateway' } },
        }),
      },
      {
        response: createAxiosResponse({
          data: { value: [{ id: 2 }] },
        }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      sleep: sleep.sleep,
      jitterRatio: 0,
      maxBatchRetries: 2,
      initialBackoffMs: 10,
      maxBackoffMs: 1000,
    })

    const out = await client.batch([{ id: '1', url: '/users' }])

    expect(calls).toHaveLength(3)
    expect(sleep.calls).toEqual([10])
    expect(out.responses['1'].body.value).toEqual([{ id: 1 }, { id: 2 }])
  })

  test('pagination throws when nextLink fetch returns non-JSON', async () => {
    const { axios } = createMockAxios([
      {
        response: createAxiosResponse({
          data: {
            responses: [
              {
                id: '1',
                status: 200,
                headers: {},
                body: {
                  value: [{ id: 1 }],
                  '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc',
                },
              },
            ],
          },
        }),
      },
      {
        response: createAxiosResponse({
          data: 'not-json',
        }),
      },
    ])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't', maxBatchRetries: 0 })

    await expect(client.batch([{ id: '1', url: '/users' }], { mode: 'strict' })).rejects.toThrow(
      /Pagination returned non-JSON/
    )
  })

  test('partial mode: pagination non-JSON does not throw (keeps nextLink)', async () => {
    const { axios, calls } = createMockAxios([
      {
        response: createAxiosResponse({
          data: {
            responses: [
              {
                id: '1',
                status: 200,
                headers: {},
                body: {
                  value: [{ id: 1 }],
                  '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc',
                },
              },
            ],
          },
        }),
      },
      {
        response: createAxiosResponse({
          data: 'not-json',
        }),
      },
    ])

    const client = new M365GraphBatchClient({ axios, getAccessToken: async () => 't', maxBatchRetries: 0 })

    const out = await client.batch([{ id: '1', url: '/users' }], { mode: 'partial' })

    expect(calls).toHaveLength(2)
    expect(out.partial).toBe(true)
    expect(out.errors).toHaveLength(1)
    expect(out.errors[0].id).toBe('1')
    expect(out.errors[0].stage).toBe('pagination')
    expect(out.responses['1'].body.value).toEqual([{ id: 1 }])
    expect(out.responses['1'].body['@odata.nextLink']).toBe('https://graph.microsoft.com/v1.0/users?$skiptoken=abc')
  })

  test('partial mode: pagination max pages does not throw (keeps nextLink)', async () => {
    const { axios, calls } = createMockAxios([
      {
        response: createAxiosResponse({
          data: {
            responses: [
              {
                id: '1',
                status: 200,
                headers: {},
                body: {
                  value: [{ id: 1 }],
                  '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc',
                },
              },
            ],
          },
        }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      maxBatchRetries: 0,
      maxPaginationPages: 0,
    })

    const out = await client.batch([{ id: '1', url: '/users' }], { mode: 'partial' })

    expect(calls).toHaveLength(1)
    expect(out.partial).toBe(true)
    expect(out.errors).toHaveLength(1)
    expect(out.errors[0].id).toBe('1')
    expect(out.errors[0].stage).toBe('pagination')
    expect(out.responses['1'].body.value).toEqual([{ id: 1 }])
    expect(out.responses['1'].body['@odata.nextLink']).toBe('https://graph.microsoft.com/v1.0/users?$skiptoken=abc')
  })

  test('pagination resolveNextLink falls back to raw link when graphOrigin is unset', async () => {
    const { axios, calls } = createMockAxios([
      {
        response: createAxiosResponse({
          data: {
            responses: [
              {
                id: '1',
                status: 200,
                headers: {},
                body: {
                  value: [{ id: 1 }],
                  '@odata.nextLink': '/users?$skiptoken=abc',
                },
              },
            ],
          },
        }),
      },
      {
        // With invalid graphBaseUrl, graphOrigin is null so resolveNextLink returns link as-is.
        matcher: (config) => config.url === 'not-a-url/users?$skiptoken=abc',
        response: createAxiosResponse({ data: { value: [{ id: 2 }] } }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      graphBaseUrl: 'not-a-url',
      maxBatchRetries: 0,
    })

    const out = await client.batch([{ id: '1', url: '/users' }])
    expect(calls).toHaveLength(2)
    expect(out.responses['1'].body.value).toEqual([{ id: 1 }, { id: 2 }])
  })

  test('jitter backoff stays deterministic when rng is injected', async () => {
    const sleep = createMockSleep()

    const { axios } = createMockAxios([
      { throw: new Error('ECONNRESET') },
      {
        response: createAxiosResponse({
          data: { responses: [{ id: '1', status: 200, headers: {}, body: {} }] },
        }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      sleep: sleep.sleep,
      jitterRatio: 0.25,
      maxBatchRetries: 5,
      initialBackoffMs: 100,
      maxBackoffMs: 1000,
      rng: () => 0, // pick minimum jitter
    })

    await client.batch([{ id: '1', url: '/x' }])

    // With jitterRatio=0.25 and rng()=0 -> min = 100*(1-0.25)=75
    expect(sleep.calls).toEqual([75])
  })

  test('pagination stops after maxPaginationPages', async () => {
    const { axios, calls } = createMockAxios([
      {
        response: createAxiosResponse({
          data: {
            responses: [
              {
                id: '1',
                status: 200,
                headers: {},
                body: {
                  value: [{ id: 1 }],
                  '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc',
                },
              },
            ],
          },
        }),
      },
      {
        response: createAxiosResponse({
          data: {
            value: [{ id: 2 }],
            '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=def',
          },
        }),
      },
    ])

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      maxPaginationPages: 1,
    })

    await expect(client.batch([{ id: '1', url: '/users' }], { mode: 'strict' })).rejects.toThrow(
      /Pagination exceeded max pages/
    )
    expect(calls).toHaveLength(2)
  })
})
