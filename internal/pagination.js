const { PaginationExceededMaxPagesError, PaginationNonJsonError } = require('../errors');

function createPaginationHandler({ getWithGlobalRetry, graphOrigin, maxPaginationPages }) {
  const resolveNextLink = (link) => {
    try {
      // Absolute
      return new URL(link).toString();
    } catch {
      if (graphOrigin) {
        return new URL(link, graphOrigin).toString();
      }
      return link;
    }
  };

  return {
    async paginateResponsesInPlace(responseList, requestMetaById, options = {}) {
      const mode = options.mode ?? 'strict';
      const onError = typeof options.onError === 'function' ? options.onError : () => {};

      // Only handles JSON bodies with { value: [], "@odata.nextLink": "..." } for GET requests.
      for (const resp of responseList) {
        const meta = requestMetaById ? requestMetaById[String(resp.id)] : null;
        if (!meta || meta.method !== 'GET') continue;
        if (!resp || resp.status < 200 || resp.status >= 300) continue;
        if (!resp.body || typeof resp.body !== 'object') continue;
        if (!Array.isArray(resp.body.value)) continue;

        const nextLink = resp.body['@odata.nextLink'];
        if (!nextLink) continue;

        const aggregated = resp.body.value.slice();

        let url = resolveNextLink(nextLink);
        let pageCount = 0;

        try {
          while (url) {
            pageCount += 1;
            if (pageCount > maxPaginationPages) {
              throw new PaginationExceededMaxPagesError({ max: maxPaginationPages, id: resp.id });
            }

            const page = await getWithGlobalRetry(url);
            if (!page || typeof page !== 'object') {
              throw new PaginationNonJsonError({ id: resp.id });
            }

            if (Array.isArray(page.value)) aggregated.push(...page.value);
            url = page['@odata.nextLink'] ? resolveNextLink(page['@odata.nextLink']) : null;
          }

          resp.body.value = aggregated;
          delete resp.body['@odata.nextLink'];
        } catch (err) {
          if (mode !== 'partial') throw err;

          resp.body.value = aggregated;
          resp.body['@odata.nextLink'] = url;

          onError(err, { id: resp.id, nextLink: url });
        }
      }
    },
  };
}

module.exports = {
  createPaginationHandler,
};
