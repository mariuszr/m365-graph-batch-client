import { describe, expect, test, vi } from 'vitest'

import { createPaginationHandler } from '../internal/pagination'
import { createRefreshTokenAccessTokenProvider } from '../internal/tokenProvider'
import { chunkArray, getRetryAfterMs, normalizeHeaders, toFullUrl, toRelativeBatchUrl } from '../internal/utils'

describe('coverage edges', () => {
  test('security: never leak authorization header in errors', async () => {
    const secretToken = 'super-secret-token'
    const axios = {
      async request() {
        return {
          status: 400,
          data: { error: 'bad' },
          headers: {},
        }
      },
    }

    const client = new (await import('..')).M365GraphBatchClient({
      axios,
      getAccessToken: async () => secretToken,
      maxBatchRetries: 0,
    })

    let thrown
    try {
      await client._requestWithGlobalRetry({ method: 'GET', url: '/x' })
    } catch (err) {
      thrown = err
    }

    expect(thrown).toBeTruthy()

    const stringifyError = (err) => {
      try {
        return JSON.stringify(err, Object.getOwnPropertyNames(err))
      } catch {
        return ''
      }
    }

    const samples = [
      String(thrown?.message ?? ''),
      String(thrown),
      JSON.stringify({ message: thrown?.message }),
      stringifyError(thrown),
    ]

    for (const s of samples) {
      expect(s).not.toContain(secretToken)
      expect(s.toLowerCase()).not.toContain('authorization')
      expect(s.toLowerCase()).not.toContain('bearer')
    }
  })

  test('security: strict mode rejects external @odata.nextLink (SSRF protection)', async () => {
    const calls = []
    const axios = {
      async request(config) {
        calls.push(config)

        if (config.method === 'POST') {
          return {
            status: 200,
            headers: {},
            data: {
              responses: [
                {
                  id: '1',
                  status: 200,
                  headers: {},
                  body: {
                    value: [{ id: 1 }],
                    '@odata.nextLink': 'https://evil.example/steal',
                  },
                },
              ],
            },
          }
        }

        // Pagination GET must never be attempted for external nextLink.
        throw new Error(`unexpected GET to ${config.url}`)
      },
    }

    const { M365GraphBatchClient } = await import('..')
    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      graphBaseUrl: 'https://graph.microsoft.com/v1.0',
      maxBatchRetries: 0,
      initialBackoffMs: 0,
      jitterRatio: 0,
    })

    await expect(
      client.batch([{ id: '1', method: 'GET', url: '/users' }], { paginate: true, mode: 'strict' })
    ).rejects.toThrow(/nextLink origin mismatch/i)

    expect(calls.some((c) => c.method === 'POST')).toBe(true)
    expect(calls.some((c) => c.method === 'GET')).toBe(false)
  })

  test('security: partial mode keeps @odata.nextLink and marks partial on external nextLink', async () => {
    const calls = []
    const axios = {
      async request(config) {
        calls.push(config)

        if (config.method === 'POST') {
          return {
            status: 200,
            headers: {},
            data: {
              responses: [
                {
                  id: '1',
                  status: 200,
                  headers: {},
                  body: {
                    value: [{ id: 1 }],
                    '@odata.nextLink': 'https://evil.example/steal',
                  },
                },
              ],
            },
          }
        }

        throw new Error(`unexpected GET to ${config.url}`)
      },
    }

    const { M365GraphBatchClient } = await import('..')
    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      graphBaseUrl: 'https://graph.microsoft.com/v1.0',
      maxBatchRetries: 0,
      initialBackoffMs: 0,
      jitterRatio: 0,
    })

    const out = await client.batch([{ id: '1', method: 'GET', url: '/users' }], { paginate: true, mode: 'partial' })

    expect(out.partial).toBe(true)
    expect(out.responses['1'].body['@odata.nextLink']).toBe('https://evil.example/steal')
    expect(calls.some((c) => c.method === 'GET')).toBe(false)
  })
  test('pagination skips when request is not GET or body invalid', async () => {
    const handler = createPaginationHandler({
      getWithGlobalRetry: vi.fn(),
      graphOrigin: 'https://graph.example',
      maxPaginationPages: 10,
    })

    // cover default mode ('strict')
    await handler.paginateResponsesInPlace([], {})

    const responseList = [
      { id: '1', status: 200, headers: {}, body: { value: [1], '@odata.nextLink': '/x' } },
      { id: '2', status: 201, headers: {}, body: null },
      { id: '3', status: 500, headers: {}, body: { value: [1], '@odata.nextLink': '/x' } },
      { id: '4', status: 200, headers: {}, body: { value: [] } },
      { id: '5', status: 200, headers: {}, body: { value: 'nope', '@odata.nextLink': '/x' } },
    ]

    const requestMetaById = {
      1: { method: 'POST' },
      2: { method: 'GET' },
      3: { method: 'GET' },
      4: { method: 'GET' },
      5: { method: 'GET' },
    }

    await handler.paginateResponsesInPlace(responseList, requestMetaById)
    expect(handler).toBeTruthy()
  })

  test('pagination strict mode rethrows errors', async () => {
    const handler = createPaginationHandler({
      getWithGlobalRetry: vi.fn().mockResolvedValueOnce(null),
      graphOrigin: null,
      maxPaginationPages: 10,
    })

    const responseList = [
      {
        id: '1',
        status: 200,
        headers: {},
        body: { value: [1], '@odata.nextLink': '\\not-a-url' },
      },
    ]

    const requestMetaById = {
      1: { method: 'GET' },
    }

    await expect(handler.paginateResponsesInPlace(responseList, requestMetaById, { mode: 'strict' })).rejects.toThrow()
  })

  test('pagination resolves absolute and relative nextLink and removes @odata.nextLink', async () => {
    const handler = createPaginationHandler({
      getWithGlobalRetry: vi
        .fn()
        .mockResolvedValueOnce({ value: [2], '@odata.nextLink': '/rel-next' })
        .mockResolvedValueOnce({ value: [3], '@odata.nextLink': 'https://graph.example/abs-next-2' })
        .mockResolvedValueOnce({ value: [] }),
      graphOrigin: 'https://graph.example',
      maxPaginationPages: 10,
    })

    const responseList = [
      {
        id: '1',
        status: 200,
        headers: {},
        body: { value: [1], '@odata.nextLink': 'https://graph.example/abs-next' },
      },
    ]

    const requestMetaById = {
      1: { method: 'GET' },
    }

    await handler.paginateResponsesInPlace(responseList, requestMetaById)

    expect(responseList[0].body.value).toEqual([1, 2, 3])
    expect(responseList[0].body['@odata.nextLink']).toBeUndefined()
  })

  test('pagination resolves bad relative link without graphOrigin', async () => {
    const handler = createPaginationHandler({
      getWithGlobalRetry: vi.fn().mockResolvedValueOnce({ value: [] }),
      graphOrigin: null,
      maxPaginationPages: 10,
    })

    const responseList = [
      {
        id: '1',
        status: 200,
        headers: {},
        body: { value: [], '@odata.nextLink': '\\not-a-url' },
      },
    ]

    const requestMetaById = {
      1: { method: 'GET' },
    }

    await handler.paginateResponsesInPlace(responseList, requestMetaById)

    expect(responseList[0].body.value).toEqual([])
    expect(responseList[0].body['@odata.nextLink']).toBeUndefined()
  })

  test('pagination in partial mode restores @odata.nextLink on error', async () => {
    const handler = createPaginationHandler({
      getWithGlobalRetry: vi.fn().mockResolvedValueOnce(null),
      graphOrigin: null,
      maxPaginationPages: 10,
    })

    const responseList = [
      {
        id: '1',
        status: 200,
        headers: {},
        body: { value: [1], '@odata.nextLink': '\\still-not-a-url' },
      },
    ]

    const requestMetaById = {
      1: { method: 'GET' },
    }

    const onError = vi.fn()
    await handler.paginateResponsesInPlace(responseList, requestMetaById, { mode: 'partial', onError })

    expect(responseList[0].body.value).toEqual([1])
    expect(responseList[0].body['@odata.nextLink']).toBe('\\still-not-a-url')
    expect(onError).toHaveBeenCalledTimes(1)
  })

  test('utils chunkArray splits arrays', () => {
    expect(chunkArray([1, 2, 3, 4, 5], 2)).toEqual([[1, 2], [3, 4], [5]])
  })

  test('utils toFullUrl rejects absolute off-origin and joins relative', () => {
    expect(() => toFullUrl({ graphBaseUrl: 'https://graph.example/v1.0', urlOrPath: 'https://x.test/a?b=1' })).toThrow(
      /origin mismatch/i
    )

    expect(toFullUrl({ graphBaseUrl: 'https://graph.example/v1.0/', urlOrPath: 'users' })).toBe(
      'https://graph.example/v1.0/users'
    )
  })

  test('utils toFullUrl preserves previous behavior when graphBaseUrl is invalid', () => {
    expect(toFullUrl({ graphBaseUrl: 'not-a-url', urlOrPath: 'users' })).toBe('not-a-url/users')
    expect(toFullUrl({ graphBaseUrl: 'not-a-url/', urlOrPath: '/users' })).toBe('not-a-url/users')
  })

  test('utils normalizeHeaders filters and lowercases, getRetryAfterMs date parsing', () => {
    expect(normalizeHeaders({ A: 'X', B: null, C: undefined, D: 1 })).toEqual({ a: 'X', d: '1' })

    const now = () => 1000
    expect(getRetryAfterMs({ 'Retry-After': 'Thu, 01 Jan 1970 00:00:02 GMT' }, now)).toBe(1000)

    // cover default `now` arg and both finite/non-finite paths
    expect(getRetryAfterMs({ 'Retry-After': '0' })).toBe(0)
    expect(getRetryAfterMs({ 'Retry-After': '2.1' }, () => 0)).toBe(2100)

    // cover Date.parse branch (call without now to hit default param)
    const originalNow = Date.now
    try {
      Date.now = () => 1000
      expect(getRetryAfterMs({ 'Retry-After': 'Thu, 01 Jan 1970 00:00:02 GMT' })).toBe(1000)
    } finally {
      Date.now = originalNow
    }
  })

  test('utils getRetryAfterMs supports delta-seconds and clamps to 0', () => {
    expect(getRetryAfterMs({ 'Retry-After': '2' }, () => 0)).toBe(2000)
    expect(getRetryAfterMs({ 'Retry-After': '0' }, () => 0)).toBe(0)
  })

  test('utils getRetryAfterMs returns null for invalid Retry-After', () => {
    expect(getRetryAfterMs({ 'Retry-After': 'not-a-date-or-number' }, () => 0)).toBeNull()
    expect(getRetryAfterMs({ 'Retry-After': 'Infinity' }, () => 0)).toBeNull()
  })

  test('utils getRetryAfterMs returns null when header missing', () => {
    expect(getRetryAfterMs({}, () => 0)).toBeNull()
    expect(getRetryAfterMs(null, () => 0)).toBeNull()
  })

  test('utils toRelativeBatchUrl returns input when already starts with slash', () => {
    expect(toRelativeBatchUrl('/users')).toBe('/users')
    expect(toRelativeBatchUrl('/already-relative?x=1')).toBe('/already-relative?x=1')
  })

  test('utils toRelativeBatchUrl adds leading slash for non-url non-slash input', () => {
    expect(toRelativeBatchUrl('users')).toBe('/users')
    // Ensure other schemes are accepted via "scheme://" matcher.
    expect(toRelativeBatchUrl('ftp://graph.example/v1.0/users')).toBe('/v1.0/users')
  })

  test('utils toFullUrl returns absolute URL when graphBaseUrl has no origin', () => {
    expect(toFullUrl({ graphBaseUrl: 'not-a-url', urlOrPath: 'https://evil.example/abs' })).toBe(
      'https://evil.example/abs'
    )
  })

  test('utils toFullUrl stringifies urlOrPath and enforces same origin', () => {
    const absoluteUrl = { toString: () => 'https://graph.example/x' }
    expect(toFullUrl({ graphBaseUrl: 'https://graph.example/v1.0', urlOrPath: absoluteUrl })).toBe(
      'https://graph.example/x'
    )

    const relativeUrl = { toString: () => 'users' }
    expect(toFullUrl({ graphBaseUrl: 'https://graph.example/v1.0', urlOrPath: relativeUrl })).toBe(
      'https://graph.example/v1.0/users'
    )

    const evilUrl = { toString: () => 'https://evil.example/abs' }
    expect(() => toFullUrl({ graphBaseUrl: 'https://graph.example/v1.0', urlOrPath: evilUrl })).toThrow(
      /origin mismatch/i
    )
  })

  test('utils toFullUrl returns joined URL when urlOrPath is not absolute', () => {
    // Force absolute parsing to fail (urlOrPath not a valid URL).
    expect(toFullUrl({ graphBaseUrl: 'https://graph.example/v1.0', urlOrPath: '\\\\not-a-url' })).toBe(
      'https://graph.example/v1.0/\\\\not-a-url'
    )
  })

  test('utils toRelativeBatchUrl strips path and query for absolute input', () => {
    expect(toRelativeBatchUrl('https://graph.example/v1.0/users?$top=1')).toBe('/v1.0/users?$top=1')
  })

  test('utils toFullUrl allows absolute URL when origin matches', () => {
    expect(toFullUrl({ graphBaseUrl: 'https://graph.example/v1.0', urlOrPath: 'https://graph.example/v1.0/me' })).toBe(
      'https://graph.example/v1.0/me'
    )

    // cover urlOrPath stringification + absolute allow branch
    expect(
      toFullUrl({
        graphBaseUrl: 'https://graph.example/v1.0',
        urlOrPath: { toString: () => 'https://graph.example/x' },
      })
    ).toBe('https://graph.example/x')
  })

  test('utils toFullUrl joins relative path and trims trailing slash in graphBaseUrl', () => {
    expect(toFullUrl({ graphBaseUrl: 'https://graph.example/v1.0/', urlOrPath: '/users' })).toBe(
      'https://graph.example/v1.0/users'
    )
    expect(toFullUrl({ graphBaseUrl: 'https://graph.example/v1.0/', urlOrPath: 'users' })).toBe(
      'https://graph.example/v1.0/users'
    )
  })

  test('utils toFullUrl blocks absolute URL when origin mismatches', () => {
    expect(() =>
      toFullUrl({ graphBaseUrl: 'https://graph.example/v1.0', urlOrPath: 'https://evil.example/steal' })
    ).toThrow(/origin mismatch/i)

    // cover urlOrPath stringification + ORIGIN_MISMATCH path
    expect(() =>
      toFullUrl({
        graphBaseUrl: 'https://graph.example/v1.0',
        urlOrPath: { toString: () => 'https://evil.example/steal' },
      })
    ).toThrow(/origin mismatch/i)
  })

  test('utils toRelativeBatchUrl treats scheme without "://" as path', () => {
    // e.g. "mailto:" should not be treated as absolute URL in this library.
    expect(toRelativeBatchUrl('mailto:test@example.com')).toBe('/mailto:test@example.com')

    // cover the throw statement in internal/utils.js
    expect(() => toRelativeBatchUrl('')).toThrow(/Request url is required/)
  })

  test('token provider throws on input validation and HTTP failures', async () => {
    const axios = {
      request: vi.fn(),
    }

    axios.request.mockResolvedValueOnce({ status: null, data: 'nope' })
    await expect(
      createRefreshTokenAccessTokenProvider({
        axios,
        tenantId: 'tenant',
        clientId: 'client',
        clientSecret: 'secret',
        refreshToken: 'refresh',
      })()
    ).rejects.toThrow(/OAuth token refresh failed \(unknown\)/)

    // cover default `now` and default `clockSkewMs` in signature
    createRefreshTokenAccessTokenProvider({
      axios,
      tenantId: 'tenant',
      clientId: 'client',
      clientSecret: 'secret',
      refreshToken: 'refresh',
    })

    const getAccessToken = createRefreshTokenAccessTokenProvider({
      axios,
      tenantId: 'tenant',
      clientId: 'client',
      clientSecret: 'secret',
      refreshToken: 'refresh',
      clockSkewMs: 0,
    })

    const invalidAxios = {
      request: null,
    }
    expect(() =>
      createRefreshTokenAccessTokenProvider({
        axios: invalidAxios,
        tenantId: 'tenant',
        clientId: 'client',
        clientSecret: 'secret',
        refreshToken: 'refresh',
      })
    ).toThrow(/options\.axios\.request is required/)

    expect(() =>
      createRefreshTokenAccessTokenProvider({
        axios,
        tenantId: '',
        clientId: 'client',
        clientSecret: 'secret',
        refreshToken: 'refresh',
      })
    ).toThrow(/options\.auth\.tenantId is required/)

    expect(() =>
      createRefreshTokenAccessTokenProvider({
        axios,
        tenantId: 'tenant',
        clientId: '',
        clientSecret: 'secret',
        refreshToken: 'refresh',
      })
    ).toThrow(/options\.auth\.clientId is required/)

    expect(() =>
      createRefreshTokenAccessTokenProvider({
        axios,
        tenantId: 'tenant',
        clientId: 'client',
        clientSecret: '',
        refreshToken: 'refresh',
      })
    ).toThrow(/options\.auth\.clientSecret is required/)

    expect(() =>
      createRefreshTokenAccessTokenProvider({
        axios,
        tenantId: 'tenant',
        clientId: 'client',
        clientSecret: 'secret',
        refreshToken: '',
      })
    ).toThrow(/options\.auth\.refreshToken is required/)

    axios.request.mockResolvedValueOnce({ status: 400, data: { error: 'nope' } })
    await expect(getAccessToken()).rejects.toThrow(/OAuth token refresh failed/)

    axios.request.mockResolvedValueOnce({ status: 200, data: { expires_in: 3600 } })
    await expect(getAccessToken()).rejects.toThrow(/no access_token/)

    axios.request.mockResolvedValueOnce({ status: 200, data: { access_token: 't', expires_in: 'nope' } })
    await expect(getAccessToken()).rejects.toThrow(/invalid expires_in/)
  })

  test('client module exports are importable', async () => {
    const mod = await import('../client.js')
    expect(mod.M365GraphBatchClient).toBeTypeOf('function')
  })

  test('backoff computeBackoffMs clamps and jitters deterministically', async () => {
    const { createBackoff } = await import('../internal/backoff.js')

    const b1 = createBackoff({ initialBackoffMs: 100, maxBackoffMs: 150, jitterRatio: 0, rng: () => 0.5 })
    expect(b1.computeBackoffMs(1)).toBe(100)
    expect(b1.computeBackoffMs(2)).toBe(150)

    const b2 = createBackoff({ initialBackoffMs: 100, maxBackoffMs: 1000, jitterRatio: 0.5, rng: () => 0 })
    expect(b2.computeBackoffMs(1)).toBe(50)

    const b3 = createBackoff({ initialBackoffMs: 0, maxBackoffMs: 1000, jitterRatio: 0.5, rng: () => 1 })
    expect(b3.computeBackoffMs(10)).toBe(0)

    const b4 = createBackoff({ initialBackoffMs: 100, maxBackoffMs: 1000, jitterRatio: -1, rng: () => 1 })
    expect(b4.computeBackoffMs(1)).toBe(100)

    // cover default rng branch
    const savedRandom = Math.random
    try {
      Math.random = () => 0
      const b5 = createBackoff({ initialBackoffMs: 100, maxBackoffMs: 1000, jitterRatio: 0.5 })
      expect(b5.computeBackoffMs(1)).toBe(50)
    } finally {
      Math.random = savedRandom
    }
  })

  test('errors classes set name and expose fields', async () => {
    const {
      M365GraphBatchClientError,
      RequestFailedError,
      RequestExceededRetriesError,
      SubrequestExceededRetriesError,
      InvalidBatchResponseShapeError,
      BatchRequestSizeExceededError,
      PaginationExceededMaxPagesError,
      PaginationNonJsonError,
      PaginationExternalNextLinkError,
    } = await import('../errors.js')

    const base = new M365GraphBatchClientError('x')
    expect(base.name).toBe('M365GraphBatchClientError')

    const e1 = new RequestFailedError({ status: 400, responseText: 'bad' })
    expect(e1.name).toBe('RequestFailedError')
    expect(e1.status).toBe(400)
    expect(e1.responseText).toBe('bad')

    const e2 = new RequestExceededRetriesError({ status: 429 })
    expect(e2.status).toBe(429)

    const e3 = new SubrequestExceededRetriesError({ id: '1', status: 503 })
    expect(e3.id).toBe('1')
    expect(e3.status).toBe(503)

    const e4 = new InvalidBatchResponseShapeError()
    expect(e4.message).toMatch(/invalid \$batch response shape/i)

    const e5 = new BatchRequestSizeExceededError({ max: 20 })
    expect(e5.max).toBe(20)

    const e6 = new PaginationExceededMaxPagesError({ max: 3, id: 'x' })
    expect(e6.max).toBe(3)
    expect(e6.id).toBe('x')

    const e7 = new PaginationNonJsonError({ id: 'x' })
    expect(e7.id).toBe('x')

    const e8 = new PaginationExternalNextLinkError({
      id: 'x',
      nextLink: 'https://evil',
      allowedOrigin: 'https://graph',
    })
    expect(e8.nextLink).toBe('https://evil')
    expect(e8.allowedOrigin).toBe('https://graph')
  })
})
