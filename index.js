const client = require('./client')
const errors = require('./errors')

module.exports = {
  ...client,
  ...errors,
}
