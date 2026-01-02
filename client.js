const {
  normalizeHeaders,
  getRetryAfterMs,
  createDefaultSleep,
  chunkArray,
  toRelativeBatchUrl,
  toFullUrl,
} = require('./internal/utils')

/**
 * @typedef {'strict'|'partial'} BatchMode
 */

/**
 * @typedef {Object} BatchRequest
 * @property {string|number} id
 * @property {string} [method]
 * @property {string} url
 * @property {Object<string,string>} [headers]
 * @property {any} [body]
 */

/**
 * A single subresponse from Graph $batch.
 * @typedef {Object} BatchSubresponse
 * @property {string} id
 * @property {number} status
 * @property {Object<string,string>} headers
 * @property {any} body
 */

/**
 * @typedef {'subrequest'|'pagination'|'auth'|'batch'} BatchErrorStage
 */

/**
 * Error reported in `mode: 'partial'`.
 *
 * - `stage: 'subrequest'`: a single subrequest exhausted retries.
 * - `stage: 'pagination'`: auto-pagination failed for a GET response that had `@odata.nextLink`.
 * - `stage: 'auth'`: token acquisition failed (offline/network) before $batch could run.
 * - `stage: 'batch'`: $batch call failed (offline/network).
 *
 * @typedef {Object} BatchPartialError
 * @property {BatchErrorStage} stage
 * @property {string} [id]
 * @property {string} type
 * @property {string} message
 * @property {number|string} [status]
 * @property {string} [code]
 * @property {number} [errno]
 * @property {string} [syscall]
 * @property {string} [hostname]
 * @property {string} [url]
 */

/**
 * @typedef {Object} BatchResultStrict
 * @property {Record<string, BatchSubresponse>} responses
 * @property {BatchSubresponse[]} responseList
 */

/**
 * @typedef {Object} BatchResultPartial
 * @property {Record<string, BatchSubresponse>} responses
 * @property {BatchSubresponse[]} responseList
 * @property {boolean} partial
 * @property {BatchPartialError[]} errors
 */

const { createBackoff } = require('./internal/backoff')
const { createPaginationHandler } = require('./internal/pagination')
const { createRefreshTokenAccessTokenProvider } = require('./internal/tokenProvider')

const {
  RequestFailedError,
  RequestExceededRetriesError,
  SubrequestExceededRetriesError,
  InvalidBatchResponseShapeError,
  BatchRequestSizeExceededError,
} = require('./errors')

class M365GraphBatchClient {
  constructor(options) {
    if (!options) throw new Error('options is required')

    const axiosInstance = options.axios || options.axiosInstance
    if (axiosInstance) {
      if (typeof axiosInstance.request !== 'function') throw new Error('options.axios.request is required')
      this._axios = axiosInstance
    } else {
      // Lazy-require, so tests can inject a mock without deps.
      let axios
      try {
        axios = require('axios')
      } catch {
        throw new Error('options.axios is required (axios dependency not found)')
      }

      this._axios = axios
    }

    this._sleep = options.sleep || createDefaultSleep()
    this._now = options.now || (() => Date.now())

    if (options.getAccessToken && typeof options.getAccessToken === 'function') {
      this._getAccessToken = options.getAccessToken
    } else if (options.auth && typeof options.auth === 'object') {
      this._getAccessToken = createRefreshTokenAccessTokenProvider({
        axios: this._axios,
        tenantId: options.auth.tenantId,
        clientId: options.auth.clientId,
        clientSecret: options.auth.clientSecret,
        refreshToken: options.auth.refreshToken,
        scope: options.auth.scope,
        now: this._now,
        clockSkewMs: options.auth.clockSkewMs,
      })
    } else {
      throw new Error('options.getAccessToken is required')
    }

    this._graphBaseUrl = options.graphBaseUrl || 'https://graph.microsoft.com/v1.0'
    this._batchPath = options.batchPath || '/$batch'

    this._graphOrigin = null
    try {
      this._graphOrigin = new URL(this._graphBaseUrl).origin
    } catch {
      // ignore
    }

    this._maxSubrequestRetries = options.maxSubrequestRetries ?? 5
    this._maxBatchRetries = options.maxBatchRetries ?? 5

    this._maxRequestsPerBatch = options.maxRequestsPerBatch ?? 20
    this._initialBackoffMs = options.initialBackoffMs ?? 250
    this._maxBackoffMs = options.maxBackoffMs ?? 30_000

    this._maxPaginationPages = options.maxPaginationPages ?? 50

    // Jitter is applied to exponential backoff to reduce thundering herd.
    this._jitterRatio = options.jitterRatio ?? 0.25

    // Allow deterministic tests.
    this._rng = options.rng

    // Default: common transient statuses for Graph.
    this._retryableStatuses = new Set(options.retryableStatuses ?? [429, 500, 502, 503, 504])

    this._backoff = createBackoff({
      initialBackoffMs: this._initialBackoffMs,
      maxBackoffMs: this._maxBackoffMs,
      jitterRatio: this._jitterRatio,
      rng: this._rng,
    })

    this._pagination = createPaginationHandler({
      getWithGlobalRetry: (url) => this._getWithGlobalRetry(url),
      graphOrigin: this._graphOrigin,
      maxPaginationPages: this._maxPaginationPages,
    })

    this._validateUrlSameOrigin = (urlOrPath) => {
      if (!this._graphOrigin) return
      try {
        const abs = new URL(String(urlOrPath))
        if (abs.origin !== this._graphOrigin) {
          const err = new Error(`Request url origin mismatch (allowed ${this._graphOrigin}): ${abs.toString()}`)
          err.code = 'ORIGIN_MISMATCH'
          throw err
        }
      } catch (err) {
        if (err?.code === 'ORIGIN_MISMATCH') throw err
        // not an absolute URL
      }
    }
  }

  /**
   * Execute Microsoft Graph batch requests.
   *
   * Default behavior is best-effort `mode: 'partial'`, meaning the method returns
   * partial results when some subrequests (or pagination) fail.
   *
   * - mode: 'strict': fail-fast; throws when a subrequest exhausts retries, when
   *   auto-pagination fails, or when the batch response is invalid.
   * - mode: 'partial' (default): returns { partial, errors } instead of throwing
   *   for retry exhaustion or pagination failures.
   *
   * Offline/network-like failures during token acquisition or during the batch call
   * itself are also returned as partial results in mode: 'partial':
   * a synthetic status: 599 response is created for each request and one global
   * entry is appended to errors[] with stage: 'auth' or stage: 'batch'.
   *
   * @param {BatchRequest[]} requests
   * @param {Object} [options]
   * @param {boolean} [options.paginate=true] Auto-paginate GET responses when nextLink is present.
   * @param {BatchMode} [options.mode='partial'] 'partial' (default) or 'strict'.
   * @returns {Promise<BatchResultStrict|BatchResultPartial>}
   */
  async batch(requests, options = {}) {
    if (!Array.isArray(requests)) throw new Error('requests must be an array')
    if (requests.length === 0) return { responses: {}, responseList: [] }

    const paginate = options.paginate ?? true
    const mode = options.mode ?? 'partial'

    const responsesById = {}
    const responseList = []

    const errors = []
    let partial = false

    const requestChunks = chunkArray(requests, this._maxRequestsPerBatch)
    for (const requestChunk of requestChunks) {
      const chunkResult = await this._executeChunkWithRetries(requestChunk, { paginate, mode })
      for (const response of chunkResult.responseList) {
        responsesById[response.id] = response
        responseList.push(response)
      }

      if (mode === 'partial') {
        partial = partial || chunkResult.partial
        errors.push(...chunkResult.errors)
      }
    }

    if (mode === 'partial') return { responses: responsesById, responseList, partial, errors }

    return { responses: responsesById, responseList }
  }

  _isRetryableStatus(status) {
    return this._retryableStatuses.has(status)
  }

  _computeBackoffMs(attempt) {
    return this._backoff.computeBackoffMs(attempt)
  }

  async _executeChunkWithRetries(requestChunk, { paginate, mode }) {
    const requestMetaById = {}
    for (const req of requestChunk) {
      const id = String(req.id)
      const method = (req.method || 'GET').toUpperCase()
      requestMetaById[id] = { method }
    }

    const errors = []
    let partial = false

    const responsesById = {}

    const classifyGlobalErrorStage = (err) => {
      const msg = typeof err?.message === 'string' ? err.message : ''
      const url = typeof err?.config?.url === 'string' ? err.config.url : ''

      if (url.includes('login.microsoftonline.com') || msg.includes('login.microsoftonline.com')) return 'auth'
      if (msg.startsWith('OAuth token refresh')) return 'auth'

      return 'batch'
    }

    const formatGlobalError = (err, stage) => ({
      stage,
      type: err?.name || 'Error',
      message: err?.message || String(err),
      code: err?.code,
      errno: err?.errno,
      syscall: err?.syscall,
      hostname: err?.hostname,
      url: err?.config?.url,
    })

    const isOfflineLikeError = (err) => {
      const msg = typeof err?.message === 'string' ? err.message : ''
      const code = err?.code
      const syscall = err?.syscall

      // Node DNS lookup failure
      if (syscall === 'getaddrinfo') return true
      if (msg.includes('getaddrinfo ENOTFOUND')) return true

      return [
        'ENOTFOUND',
        'EAI_AGAIN',
        'ECONNREFUSED',
        'ETIMEDOUT',
        'ECONNRESET',
        'ENETUNREACH',
        'EHOSTUNREACH',
      ].includes(code)
    }

    const ensureSyntheticBatchFailureResponses = (stage, message) => {
      for (const req of requestChunk) {
        if (responsesById[req.id]) continue
        responsesById[req.id] = {
          id: String(req.id),
          status: 599,
          headers: {},
          body: {
            error: {
              code: 'BatchRequestFailed',
              message,
            },
            stage,
          },
        }
      }
    }

    const ensureSyntheticSubrequestFailureResponse = (id, errorCode, message, status) => {
      responsesById[id] = {
        id: String(id),
        status: 599,
        headers: {},
        body: {
          error: {
            code: errorCode,
            message,
          },
          status,
        },
      }
    }

    // Preflight: if any subrequest has an off-origin absolute URL, treat it as a partial subrequest error.
    const offOrigin = []
    if (this._graphOrigin) {
      for (const req of requestChunk) {
        try {
          this._validateUrlSameOrigin(req.url)
        } catch (err) {
          if (err?.code === 'ORIGIN_MISMATCH') offOrigin.push(req)
          else throw err
        }
      }
    }

    let effectiveChunk = requestChunk

    if (offOrigin.length > 0) {
      const message = `Request url origin mismatch (allowed ${this._graphOrigin}): ${offOrigin[0].url}`
      if (mode !== 'partial') throw new Error(message)

      partial = true
      for (const req of offOrigin) {
        ensureSyntheticSubrequestFailureResponse(req.id, 'ORIGIN_MISMATCH', message, 599)
        errors.push({
          id: String(req.id),
          stage: 'subrequest',
          type: 'OriginMismatchError',
          message,
          code: 'ORIGIN_MISMATCH',
          url: String(req.url),
        })
      }

      effectiveChunk = requestChunk.filter((r) => !offOrigin.some((e) => String(e.id) === String(r.id)))
    }

    // First, execute the whole chunk once. Then, isolate retryable subresponses.
    let initial
    if (effectiveChunk.length === 0) {
      const ordered = requestChunk.map((r) => responsesById[r.id]).filter(Boolean)
      return mode === 'partial'
        ? { responsesById, responseList: ordered, partial, errors }
        : { responsesById, responseList: ordered }
    }

    try {
      initial = await this._postBatchWithGlobalRetry(effectiveChunk)
    } catch (err) {
      // In partial mode, only swallow offline/network failures.
      // Other failures (401, invalid $batch shape, invalid_grant, etc.) still throw.
      if (mode !== 'partial' || !isOfflineLikeError(err)) throw err

      const stage = classifyGlobalErrorStage(err)
      partial = true
      errors.push(formatGlobalError(err, stage))
      ensureSyntheticBatchFailureResponses(stage, errors[errors.length - 1].message)

      const ordered = requestChunk.map((r) => responsesById[r.id]).filter(Boolean)

      return { responsesById, responseList: ordered, partial, errors }
    }

    for (const r of initial.responses) {
      responsesById[r.id] = r
    }

    const retryState = new Map()
    const getAttempts = (id) => retryState.get(id) ?? 0
    const incAttempts = (id) => retryState.set(id, getAttempts(id) + 1)

    let pending = effectiveChunk.slice()
    // Apply initial results.
    pending = pending.filter((req) => {
      const response = responsesById[req.id]
      // Missing response should be treated as retryable (defensive).
      if (!response) return true
      return this._isRetryableStatus(response.status)
    })

    // If any retryable subresponses exist, retry only those.
    while (pending.length > 0) {
      const exhausted = []
      const retryList = []
      for (const req of pending) {
        const nextAttempts = getAttempts(req.id) + 1
        if (nextAttempts > this._maxSubrequestRetries) {
          exhausted.push(req)
        } else {
          retryList.push(req)
        }
      }

      if (exhausted.length > 0) {
        if (mode !== 'partial') {
          const req = exhausted[0]
          const lastResponse = responsesById[req.id]
          const status = lastResponse ? lastResponse.status : 'unknown'
          throw new SubrequestExceededRetriesError({ id: req.id, status })
        }

        partial = true
        for (const req of exhausted) {
          const lastResponse = responsesById[req.id]
          const status = lastResponse ? lastResponse.status : 'unknown'
          const err = new SubrequestExceededRetriesError({ id: req.id, status })
          errors.push({
            id: String(req.id),
            stage: 'subrequest',
            type: err.name,
            message: err.message,
            status,
          })

          if (!lastResponse)
            ensureSyntheticSubrequestFailureResponse(req.id, 'SubrequestExceededRetries', err.message, status)
        }
      }

      if (retryList.length === 0) break

      // Calculate delay: prefer per-subrequest Retry-After.
      let delayMs = null
      for (const req of retryList) {
        const response = responsesById[req.id]
        const ra = response ? getRetryAfterMs(response.headers, this._now) : null
        if (ra !== null) delayMs = delayMs === null ? ra : Math.max(delayMs, ra)
      }

      // If no Retry-After headers, do exponential backoff based on max attempts.
      if (delayMs === null) {
        const maxAttempts = Math.max(0, ...retryList.map((r) => getAttempts(r.id)))
        delayMs = this._computeBackoffMs(maxAttempts + 1)
      }

      if (delayMs > 0) await this._sleep(delayMs)

      for (const req of retryList) {
        incAttempts(req.id)
      }

      const retryBatch = await this._postBatchWithGlobalRetry(retryList)
      for (const r of retryBatch.responses) {
        responsesById[r.id] = r
      }

      pending = retryList.filter((req) => {
        const resp = responsesById[req.id]
        if (!resp) return true
        return this._isRetryableStatus(resp.status)
      })
    }

    const responseList = Object.values(responsesById)

    if (paginate) {
      await this._paginateResponsesInPlace(responseList, requestMetaById, {
        mode,
        onError: (err, ctx) => {
          partial = true
          errors.push({
            id: String(ctx.id),
            stage: 'pagination',
            type: err.name,
            message: err.message,
          })
        },
      })
    }

    // Preserve stable order of the original requestChunk.
    const ordered = requestChunk.map((r) => responsesById[r.id]).filter(Boolean)

    if (mode === 'partial') return { responsesById, responseList: ordered, partial, errors }

    return { responsesById, responseList: ordered }
  }

  async _paginateResponsesInPlace(responseList, requestMetaById, options) {
    return this._pagination.paginateResponsesInPlace(responseList, requestMetaById, options)
  }

  async _getWithGlobalRetry(urlOrPath) {
    const req = { method: 'GET', url: urlOrPath }
    return this._requestWithGlobalRetry(req)
  }

  async _requestWithGlobalRetry({ method, url, headers, body }) {
    this._validateUrlSameOrigin(url)

    let attempt = 0

    while (true) {
      const token = await this._getAccessToken()
      const fullUrl = this._toFullUrl(url)

      let response
      try {
        response = await this._axios.request({
          url: fullUrl,
          method,
          headers: {
            authorization: `Bearer ${token}`,
            ...(body ? { 'content-type': 'application/json' } : {}),
            ...headers,
          },
          data: body,
          // We do retry handling ourselves.
          validateStatus: () => true,
        })
      } catch (err) {
        attempt += 1
        if (attempt > this._maxBatchRetries) throw err
        const backoffMs = this._computeBackoffMs(attempt)
        if (backoffMs > 0) await this._sleep(backoffMs)
        continue
      }

      const status = response.status
      const responseHeaders = normalizeHeaders(response.headers)

      if (status >= 200 && status < 300) {
        return response.data ?? null
      }

      if (!this._isRetryableStatus(status)) {
        const responseText = typeof response.data === 'string' ? response.data : JSON.stringify(response.data ?? '')
        throw new RequestFailedError({ status, responseText })
      }

      attempt += 1
      if (attempt > this._maxBatchRetries) {
        throw new RequestExceededRetriesError({ status })
      }

      const retryAfterMs = getRetryAfterMs(responseHeaders, this._now)
      const backoffMs = this._computeBackoffMs(attempt)
      const delayMs = retryAfterMs ?? backoffMs
      if (delayMs > 0) await this._sleep(delayMs)
    }
  }

  async _postBatchWithGlobalRetry(requestChunk) {
    for (const req of requestChunk) {
      try {
        this._validateUrlSameOrigin(req.url)
      } catch (err) {
        if (err?.code === 'ORIGIN_MISMATCH') {
          err.stage = 'subrequest'
        }
        throw err
      }
    }

    if (requestChunk.length > this._maxRequestsPerBatch) {
      throw new BatchRequestSizeExceededError({ max: this._maxRequestsPerBatch })
    }

    const payload = {
      requests: requestChunk.map((r) => ({
        id: String(r.id),
        method: (r.method || 'GET').toUpperCase(),
        url: toRelativeBatchUrl(r.url),
        headers: r.headers || undefined,
        body: r.body || undefined,
      })),
    }

    const result = await this._requestWithGlobalRetry({
      method: 'POST',
      url: this._batchPath,
      body: payload,
    })

    if (!result || !Array.isArray(result.responses)) {
      throw new InvalidBatchResponseShapeError()
    }

    // Normalize headers for downstream Retry-After parsing.
    const responses = result.responses.map((r) => {
      const id = String(r.id)
      const status = r.status
      const headers = normalizeHeaders(r.headers)
      const body = r.body
      return { id, status, headers, body }
    })

    const out = { responses }
    return out
  }

  _toFullUrl(urlOrPath) {
    const args = { graphBaseUrl: this._graphBaseUrl, urlOrPath }
    return toFullUrl(args)
  }
}

module.exports = {
  M365GraphBatchClient,
  getRetryAfterMs,
  normalizeHeaders,
  toRelativeBatchUrl,
}
