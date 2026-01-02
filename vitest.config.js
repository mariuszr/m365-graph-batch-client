module.exports = {
  test: {
    environment: 'node',
    include: ['test/**/*.test.{js,mjs}'],
    coverage: {
      provider: 'v8',
      reporter: ['text', 'lcov', 'json'],
      reportsDirectory: 'coverage',
      all: true,
      include: ['client.js', 'errors.js', 'index.js', 'internal/**/*.js'],
      exclude: ['**/node_modules/**', 'test/**', '.github/**', 'coverage/**'],
    },
  },
}
