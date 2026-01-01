# m365-graph-batch-client

Resilient Microsoft Graph `$batch` executor for Node.js.

![CI](https://github.com/mariuszr/m365-graph-batch-client/actions/workflows/ci.yml/badge.svg)
[![codecov](https://codecov.io/gh/mariuszr/m365-graph-batch-client/branch/main/graph/badge.svg)](https://codecov.io/gh/mariuszr/m365-graph-batch-client)
![License](https://img.shields.io/badge/license-MIT-green)
![Node](https://img.shields.io/badge/node-%3E%3D18-blue)

A small, pragmatic helper for running Microsoft Graph `$batch` requests without the usual footguns:
retries with backoff, best-effort partial results, and optional auto-pagination.

## Install

```bash
npm i m365-graph-batch-client
```

## Motivation

Graph `$batch` is great for reducing round-trips, but real-world usage usually needs a bit more than “just send the request”:
throttling (429), transient 5xx, partial failures inside a batch, and `@odata.nextLink` pagination.
This library wraps those concerns so your application code can stay simple.

## Features

- Automatic chunking to Graph’s 20 subrequest limit
- Retries for common transient statuses (429/5xx) with exponential backoff + jitter
- Honors `Retry-After` when present
- `mode: 'partial'` (default) returns as much as possible + an `errors[]` list
- `mode: 'strict'` throws on failures
- Optional auto-pagination for successful `GET` responses

## Non-goals

- Not a full Microsoft Graph SDK
- Not a dependency graph runner for batch subrequests
- Pagination is only for JSON bodies that look like `{ value: [] }`

## Quick Start

```js
const { M365GraphBatchClient } = require('m365-graph-batch-client');

const client = new M365GraphBatchClient({
  // Provide your own token getter (recommended)
  getAccessToken: async () => process.env.MS_GRAPH_TOKEN,
});

const result = await client.batch(
  [
    { id: '1', url: '/users?$top=1' },
    { id: '2', url: '/groups?$top=1' },
  ],
  {
    mode: 'partial',
    paginate: true,
  }
);

console.log(result.responses['1'].status, result.responses['1'].body);
if (result.partial) console.warn('Some requests failed', result.errors);
```

## Examples

### Auth examples

#### Provide your own access token getter (recommended)

```js
const { M365GraphBatchClient } = require('m365-graph-batch-client');

const client = new M365GraphBatchClient({
  getAccessToken: async () => process.env.MS_GRAPH_TOKEN,
});
```

#### Use refresh-token auth (convenience)

Note: your app must be granted appropriate Microsoft Graph permissions. The default scope is
`https://graph.microsoft.com/.default offline_access`.

```js
const { M365GraphBatchClient } = require('m365-graph-batch-client');

const client = new M365GraphBatchClient({
  auth: {
    tenantId: process.env.MS_TENANT_ID,
    clientId: process.env.MS_CLIENT_ID,
    clientSecret: process.env.MS_CLIENT_SECRET,
    refreshToken: process.env.MS_REFRESH_TOKEN,
    // scope: 'https://graph.microsoft.com/.default offline_access',
  },
});
```

### Strict vs partial

```js
// Strict: throws when something fails
await client.batch(requests, { mode: 'strict' });

// Partial (default): returns partial results + errors[]
const out = await client.batch(requests, { mode: 'partial' });
if (out.partial) console.warn(out.errors);
```

### Inject your own axios instance

```js
const axios = require('axios');

const client = new M365GraphBatchClient({
  axios,
  getAccessToken: async () => process.env.MS_GRAPH_TOKEN,
});
```

## Notes

- Graph `$batch` supports up to 20 subrequests per call; this library chunks automatically (configurable via `maxRequestsPerBatch`).
- When `paginate: true`, only successful `GET` responses with `{ value: [] }` bodies are auto-paginated.

## API

### `new M365GraphBatchClient(options)`

- `getAccessToken: async () => string` (recommended) Returns a valid Microsoft Graph access token.
- `auth: { tenantId, clientId, clientSecret, refreshToken, scope? }` Convenience option: the client can fetch tokens using a refresh token.
- `graphBaseUrl?: string` Defaults to `https://graph.microsoft.com/v1.0`.
- `batchPath?: string` Defaults to `/$batch`.
- Retry/backoff tuning: `maxSubrequestRetries`, `maxBatchRetries`, `initialBackoffMs`, `maxBackoffMs`, `jitterRatio`.
- Limits: `maxRequestsPerBatch` (default 20), `maxPaginationPages` (default 50).

### `await client.batch(requests, options?)`

- `requests`: array of `{ id, method?, url, headers?, body? }`.
- `options.mode`: `partial` (default) or `strict`.
- `options.paginate`: `true` (default) will auto-follow `@odata.nextLink` for successful `GET` responses.

#### Return value

- In `partial` mode (default): `{ responses, responseList, partial, errors }`.
- In `strict` mode: `{ responses, responseList }` and the method throws on failures.

`responses` is a map keyed by request `id`, each entry looks like:
`{ id, status, headers, body }`.

`errors` items include `stage` (`subrequest`, `pagination`, `auth`, `batch`) with a human-readable `message`.

## FAQ

### Where do I get an access token?

This library expects you to provide a valid Microsoft Graph access token.
How you acquire it depends on your app (device code, client credentials, on-behalf-of, etc.).

### Why is `partial` mode the default?

In production, you often want “best effort” behavior: successful subresponses are still useful,
and you can decide what to do with the failures.

### What is HTTP status `599`?

In `mode: 'partial'`, offline/network-like failures during token acquisition or the `$batch` call
are represented as synthetic `599` subresponses so you still get a complete `responses` map.

## Codecov setup

1. Connect the repo in Codecov.
2. Add `CODECOV_TOKEN` as a GitHub Actions secret.
3. Push to `main` (or open a PR) to upload coverage.

## Development

- Run tests: `npm test`
- Run coverage: `npm run coverage`

## Contributing

Issues and PRs are welcome. Please include tests for behavioral changes.

## License

License: MIT
