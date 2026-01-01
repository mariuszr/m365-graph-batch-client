# m365-graph-batch-client

A small, pragmatic Node.js helper for running Microsoft Graph `$batch` requests without the usual footguns.
It focuses on being resilient by default: retries with backoff, best-effort partial results, and optional auto-pagination.

Resilient Microsoft Graph batch executor for Node.js (retry/backoff, partial results, auto-pagination)

## Install

```bash
npm i m365-graph-batch-client
```

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

## Notes

- `$batch` supports up to 20 subrequests per call; this library chunks automatically (configurable via `maxRequestsPerBatch`).
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


## License

License: MIT
