const { M365GraphBatchClient } = require('..');

function createAxiosResponse({ status = 200, data, headers = {} }) {
  const normalizedHeaders = {};
  for (const [k, v] of Object.entries(headers || {})) normalizedHeaders[String(k).toLowerCase()] = String(v);

  return {
    status,
    data,
    headers: normalizedHeaders,
  };
}

function createMockSleep() {
  const calls = [];
  const sleep = async (ms) => {
    calls.push(ms);
  };
  return { sleep, calls };
}

function createMockAxios(sequence) {
  const calls = [];
  let idx = 0;

  const axios = {
    async request(config) {
      calls.push(config);

      if (idx >= sequence.length) {
        throw new Error(`Mock axios out of responses (call ${calls.length})`);
      }

      const step = sequence[idx++];
      if (step.throw) throw step.throw;
      return step.response;
    },
  };

  return { axios, calls };
}

describe('m365GraphBatchClient offline + partial mode', () => {
  test("partial mode: type falls back to 'Error' when err.name missing", async () => {
    const offlineErr = { message: 'ENOTFOUND', code: 'ENOTFOUND', config: { url: 'https://login.microsoftonline.com/x' } };

    const axios = {
      async request() {
        throw offlineErr;
      },
    };

    const client = new M365GraphBatchClient({
      axios,
      auth: {
        tenantId: 'tenant',
        clientId: 'client',
        clientSecret: 'secret',
        refreshToken: 'refresh',
      },
      maxBatchRetries: 0,
      initialBackoffMs: 0,
      jitterRatio: 0,
    });

    const out = await client.batch([{ id: '1', url: '/x' }], { mode: 'partial' });

    expect(out.partial).toBe(true);
    expect(out.errors).toHaveLength(1);
    expect(out.errors[0].type).toBe('Error');
  });

  test('partial mode: offline auth does not throw (ENOTFOUND)', async () => {
    const sleep = createMockSleep();

    const offlineErr = new Error('ENOTFOUND');
    offlineErr.name = 'AxiosError';
    offlineErr.code = 'ENOTFOUND';
    offlineErr.config = { url: 'https://login.microsoftonline.com/tenant/oauth2/v2.0/token' };

    const { axios } = createMockAxios([{ throw: offlineErr }]);

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
    });

    const out = await client.batch([{ id: '1', url: '/x' }], { mode: 'partial' });

    expect(out.partial).toBe(true);
    expect(out.errors).toHaveLength(1);
    expect(out.errors[0].stage).toBe('auth');
    expect(out.errors[0].code).toBe('ENOTFOUND');
    expect(out.responses['1'].status).toBe(599);
  });

  test('partial mode: offline auth DNS failure does not throw (syscall=getaddrinfo)', async () => {
    const sleep = createMockSleep();

    const offlineErr = new Error('getaddrinfo ENOTFOUND login.microsoftonline.com');
    offlineErr.name = 'AxiosError';
    offlineErr.code = 'ENOTFOUND';
    offlineErr.errno = -3008;
    offlineErr.syscall = 'getaddrinfo';
    offlineErr.hostname = 'login.microsoftonline.com';
    offlineErr.config = { url: 'https://login.microsoftonline.com/tenant/oauth2/v2.0/token' };

    const { axios } = createMockAxios([{ throw: offlineErr }]);

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
    });

    const out = await client.batch([{ id: '1', url: '/x' }], { mode: 'partial' });

    expect(out.partial).toBe(true);
    expect(out.errors).toHaveLength(1);
    expect(out.errors[0].stage).toBe('auth');
    expect(out.responses['1'].status).toBe(599);
  });

  test('partial mode: offline $batch network error becomes synthetic subresponses', async () => {
    const sleep = createMockSleep();

    const offlineErr = new Error('ECONNRESET');
    offlineErr.name = 'AxiosError';
    offlineErr.code = 'ECONNRESET';
    offlineErr.config = { url: 'https://graph.microsoft.com/v1.0/$batch' };

    const { axios } = createMockAxios([{ throw: offlineErr }]);

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      sleep: sleep.sleep,
      initialBackoffMs: 0,
      jitterRatio: 0,
      maxBatchRetries: 0,
    });

    const out = await client.batch([
      { id: '1', url: '/users' },
      { id: '2', url: '/groups' },
    ]);

    expect(out.partial).toBe(true);
    expect(out.errors).toHaveLength(1);
    expect(out.errors[0].stage).toBe('batch');

    expect(out.responses['1'].status).toBe(599);
    expect(out.responses['2'].status).toBe(599);
  });

  test('partial mode: duplicate request ids are skipped in synthetic creation', async () => {
    const sleep = createMockSleep();

    const offlineErr = new Error('ECONNRESET');
    offlineErr.name = 'AxiosError';
    offlineErr.code = 'ECONNRESET';
    offlineErr.config = { url: 'https://graph.microsoft.com/v1.0/$batch' };

    const { axios } = createMockAxios([{ throw: offlineErr }]);

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      sleep: sleep.sleep,
      initialBackoffMs: 0,
      jitterRatio: 0,
      maxBatchRetries: 0,
    });

    const out = await client.batch([
      { id: '1', url: '/users' },
      { id: '1', url: '/groups' },
    ]);

    expect(out.partial).toBe(true);
    expect(Object.keys(out.responses)).toEqual(['1']);
    expect(out.responses['1'].status).toBe(599);
  });

  test('partial mode: auth classification by OAuth message prefix', async () => {
    const sleep = createMockSleep();

    const offlineErr = new Error('OAuth token refresh failed: offline');
    offlineErr.name = 'AxiosError';
    offlineErr.code = 'ENOTFOUND';
    offlineErr.config = { url: 'https://example.invalid/does-not-matter' };

    const { axios } = createMockAxios([{ throw: offlineErr }]);

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
    });

    const out = await client.batch([{ id: '1', url: '/x' }], { mode: 'partial' });

    expect(out.partial).toBe(true);
    expect(out.errors).toHaveLength(1);
    expect(out.errors[0].stage).toBe('auth');
  });

  test('partial mode: handles non-string err.message and err.config.url', async () => {
    const sleep = createMockSleep();

    const offlineErr = new Error('ECONNRESET');
    offlineErr.name = 'AxiosError';
    offlineErr.code = 'ECONNRESET';
    offlineErr.message = null;
    offlineErr.config = { url: 123 };

    const { axios } = createMockAxios([{ throw: offlineErr }]);

    const client = new M365GraphBatchClient({
      axios,
      getAccessToken: async () => 't',
      sleep: sleep.sleep,
      initialBackoffMs: 0,
      jitterRatio: 0,
      maxBatchRetries: 0,
    });

    const out = await client.batch([{ id: '1', url: '/users' }]);

    expect(out.partial).toBe(true);
    expect(out.errors).toHaveLength(1);
    expect(out.errors[0].stage).toBe('batch');
  });

  test('partial mode: global error message-based auth classification', async () => {
    const sleep = createMockSleep();

    const offlineErr = new Error('getaddrinfo ENOTFOUND login.microsoftonline.com');
    offlineErr.name = 'AxiosError';
    offlineErr.code = 'ENOTFOUND';
    offlineErr.config = { url: 'https://somewhere-else.example.com' };

    const { axios } = createMockAxios([{ throw: offlineErr }]);

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
    });

    const out = await client.batch([{ id: '1', url: '/x' }], { mode: 'partial' });

    expect(out.partial).toBe(true);
    expect(out.errors).toHaveLength(1);
    expect(out.errors[0].stage).toBe('auth');
    expect(out.responses['1'].status).toBe(599);
  });

  test('partial mode: non-offline errors still throw (sanity)', async () => {
    const { axios } = createMockAxios([
      {
        response: createAxiosResponse({
          status: 400,
          data: { error: 'invalid_grant' },
        }),
      },
    ]);

    const client = new M365GraphBatchClient({
      axios,
      auth: {
        tenantId: 'tenant',
        clientId: 'client',
        clientSecret: 'secret',
        refreshToken: 'refresh',
      },
      initialBackoffMs: 0,
      jitterRatio: 0,
      maxBatchRetries: 0,
    });

    await expect(client.batch([{ id: '1', url: '/x' }], { mode: 'partial' })).rejects.toThrow(/OAuth token refresh failed/);
  });
});
