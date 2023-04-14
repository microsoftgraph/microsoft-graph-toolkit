const cfg = require('../../../.eslintrc.js');

const config = Object.assign(cfg, {
  parserOptions: {
    project: 'tsconfig.provider.json',
    sourceType: 'module'
  }
});

module.exports = config;
