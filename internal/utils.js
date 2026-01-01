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

function toRelativeBatchUrl(urlOrPath) {
  if (!urlOrPath) throw new Error('Request url is required')

  // Graph batch requires relative urls. If user passes absolute, strip origin.
  try {
    const maybeUrl = new URL(urlOrPath)
    return `${maybeUrl.pathname}${maybeUrl.search}`
  } catch {
    if (String(urlOrPath).startsWith('/')) return String(urlOrPath)
    return `/${String(urlOrPath)}`
  }
}

function toFullUrl({ graphBaseUrl, urlOrPath }) {
  // If absolute, keep it.
  try {
    return new URL(urlOrPath).toString()
  } catch {
    const base = graphBaseUrl.replace(/\/$/, '')
    const path = String(urlOrPath).startsWith('/') ? String(urlOrPath) : `/${String(urlOrPath)}`
    return `${base}${path}`
  }
}

module.exports = {
  normalizeHeaders,
  getRetryAfterMs,
  createDefaultSleep,
  chunkArray,
  toRelativeBatchUrl,
  toFullUrl,
}
