function createBackoff({ initialBackoffMs, maxBackoffMs, jitterRatio, rng }) {
  const random = rng || Math.random

  return {
    computeBackoffMs(attempt) {
      const unclamped = initialBackoffMs * 2 ** Math.max(0, attempt - 1)
      const clamped = Math.min(maxBackoffMs, unclamped)

      const effectiveJitterRatio = Math.max(0, jitterRatio)
      if (effectiveJitterRatio === 0 || clamped === 0) return clamped

      const min = clamped * (1 - effectiveJitterRatio)
      const max = clamped * (1 + effectiveJitterRatio)
      return Math.floor(min + (max - min) * random())
    },
  }
}

module.exports = {
  createBackoff,
}
