const { tailwindTransform } = require('postcss-lit');

module.exports = {
  mode: 'jit',
  content: {
    files: ['../mgt-components/**/*.ts'],
    transform: {
      ts: tailwindTransform
    }
  }
};
