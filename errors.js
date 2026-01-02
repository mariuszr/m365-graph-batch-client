class M365GraphBatchClientError extends Error {
  constructor(message) {
    super(message)
    this.name = this.constructor.name
  }
}

class RequestFailedError extends M365GraphBatchClientError {
  constructor({ status, responseText }) {
    super(`Request failed (${status}): ${responseText}`)
    this.status = status
    this.responseText = responseText
  }
}

class RequestExceededRetriesError extends M365GraphBatchClientError {
  constructor({ status }) {
    super(`Request exceeded retries (last status ${status})`)
    this.status = status
  }
}

class SubrequestExceededRetriesError extends M365GraphBatchClientError {
  constructor({ id, status }) {
    super(`Subrequest ${id} exceeded retries (last status ${status})`)
    this.id = id
    this.status = status
  }
}

class InvalidBatchResponseShapeError extends M365GraphBatchClientError {
  constructor() {
    super('Invalid $batch response shape')
  }
}

class BatchRequestSizeExceededError extends M365GraphBatchClientError {
  constructor({ max }) {
    super(`Batch request size exceeds ${max}`)
    this.max = max
  }
}

class PaginationExceededMaxPagesError extends M365GraphBatchClientError {
  constructor({ max, id }) {
    super(`Pagination exceeded max pages (${max}) for ${id}`)
    this.max = max
    this.id = id
  }
}

class PaginationNonJsonError extends M365GraphBatchClientError {
  constructor({ id }) {
    super(`Pagination returned non-JSON for ${id}`)
    this.id = id
  }
}

class PaginationExternalNextLinkError extends M365GraphBatchClientError {
  constructor({ id, nextLink, allowedOrigin }) {
    super(`Pagination nextLink origin mismatch for ${id} (allowed ${allowedOrigin}): ${nextLink}`)
    this.id = id
    this.nextLink = nextLink
    this.allowedOrigin = allowedOrigin
  }
}

module.exports = {
  M365GraphBatchClientError,
  RequestFailedError,
  RequestExceededRetriesError,
  SubrequestExceededRetriesError,
  InvalidBatchResponseShapeError,
  BatchRequestSizeExceededError,
  PaginationExceededMaxPagesError,
  PaginationNonJsonError,
  PaginationExternalNextLinkError,
}
