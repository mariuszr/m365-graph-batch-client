function normalizeHeaders(headers) {
  if (!headers) return {}

  const normalized = {}
  for (const [key, value] of Object.entries(headers)) {
    if (value === undefined || value === null) continue
    normalized[String(key).toLowerCase()] = String(value)
  }
  return normalized
}

function getRetryAfterMs(headers, now = () => Date.now()) {
  const normalized = normalizeHeaders(headers)
  const retryAfter = normalized['retry-after']
  if (!retryAfter) return null

  const seconds = Number(retryAfter)
  if (Number.isFinite(seconds)) {
    return Math.max(0, Math.floor(seconds * 1000))
  }

  const dateMs = Date.parse(retryAfter)
  if (Number.isFinite(dateMs)) {
    return Math.max(0, dateMs - now())
  }

  return null
}

function createDefaultSleep() {
  return (ms) => new Promise((resolve) => setTimeout(resolve, ms))
}

function chunkArray(arr, size) {
  const chunks = []
  for (let i = 0; i < arr.length; i += size) chunks.push(arr.slice(i, i + size))
  return chunks
}

function isAbsoluteUrl(value) {
  // Deterministic absolute URL check:
  // - requires a scheme (RFC 3986) using the "scheme://" form
  // - avoids Node's `new URL('/path')` behavior (treated as "null:" URL)
  return /^[a-zA-Z][a-zA-Z0-9+.-]*:\/\//.test(String(value))
}

function toRelativeBatchUrl(urlOrPath) {
  if (!urlOrPath) throw new Error('Request url is required')

  const str = String(urlOrPath)

  // Graph batch requires relative urls. If user passes absolute, strip origin.
  if (isAbsoluteUrl(str)) {
    const maybeUrl = new URL(str)
    return `${maybeUrl.pathname}${maybeUrl.search}`
  }

  if (str.startsWith('/')) return str
  return `/${str}`
}

function toFullUrl({ graphBaseUrl, urlOrPath }) {
  // If main base URL is invalid, preserve previous behavior.
  let graphOrigin = null
  try {
    graphOrigin = new URL(graphBaseUrl).origin
  } catch {
    graphOrigin = null
  }

  const input = String(urlOrPath)

  // If absolute, enforce same-origin to avoid SSRF.
  if (isAbsoluteUrl(input)) {
    const absolute = new URL(input)

    const originMatches = !graphOrigin || absolute.origin === graphOrigin
    if (!originMatches) {
      const err = new Error(`Request url origin mismatch (allowed ${graphOrigin}): ${absolute.toString()}`)
      err.code = 'ORIGIN_MISMATCH'
      throw err
    }

    return absolute.toString()
  }

  const base = String(graphBaseUrl).replace(/\/$/, '')
  const path = input.startsWith('/') ? input : `/${input}`
  return `${base}${path}`
}

module.exports = {
  normalizeHeaders,
  getRetryAfterMs,
  createDefaultSleep,
  chunkArray,
  toRelativeBatchUrl,
  toFullUrl,
}
