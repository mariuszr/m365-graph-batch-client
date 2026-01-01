function createRefreshTokenAccessTokenProvider({
  axios,
  tenantId,
  clientId,
  clientSecret,
  refreshToken,
  scope,
  now = () => Date.now(),
  clockSkewMs = 30_000,
}) {
  if (!axios || typeof axios.request !== 'function') throw new Error('options.axios.request is required')
  if (!tenantId) throw new Error('options.auth.tenantId is required')
  if (!clientId) throw new Error('options.auth.clientId is required')
  if (!clientSecret) throw new Error('options.auth.clientSecret is required')
  if (!refreshToken) throw new Error('options.auth.refreshToken is required')

  const effectiveScope = scope || 'https://graph.microsoft.com/.default offline_access'

  let cachedToken = null
  let cachedTokenExpiresAtMs = 0
  let pendingRefresh = null

  async function refresh() {
    const url = `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`

    const form = new URLSearchParams()
    form.set('client_id', clientId)
    form.set('client_secret', clientSecret)
    form.set('grant_type', 'refresh_token')
    form.set('refresh_token', refreshToken)
    form.set('scope', effectiveScope)

    const response = await axios.request({
      method: 'POST',
      url,
      headers: {
        'content-type': 'application/x-www-form-urlencoded',
      },
      data: form.toString(),
      validateStatus: () => true,
    })

    if (!response || response.status < 200 || response.status >= 300) {
      const bodyText = typeof response?.data === 'string' ? response.data : JSON.stringify(response?.data ?? '')
      throw new Error(`OAuth token refresh failed (${response?.status ?? 'unknown'}): ${bodyText}`)
    }

    const token = response?.data?.access_token
    if (!token) throw new Error('OAuth token refresh returned no access_token')

    const expiresInSeconds = Number(response?.data?.expires_in)
    if (!Number.isFinite(expiresInSeconds)) throw new Error('OAuth token refresh returned invalid expires_in')

    cachedToken = token
    cachedTokenExpiresAtMs = now() + Math.max(0, expiresInSeconds * 1000)

    return cachedToken
  }

  return async function getAccessToken() {
    const timeNow = now()

    if (cachedToken && cachedTokenExpiresAtMs - clockSkewMs > timeNow) {
      return cachedToken
    }

    if (!pendingRefresh) {
      pendingRefresh = refresh().finally(() => {
        pendingRefresh = null
      })
    }

    return pendingRefresh
  }
}

module.exports = {
  createRefreshTokenAccessTokenProvider,
}
