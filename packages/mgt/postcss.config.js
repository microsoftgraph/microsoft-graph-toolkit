module.exports = {
  plugins: {
    tailwindcss: {},
    autoprefixer: {},
    cssnano: {
      preset: [
        'default',
        {
          discardComments: {
            removeAll: true
          },
          minifySelectors: false,
          minifyGradients: false,
          minifyParams: false,
          normalizeWhitespace: false
        }
      ]
    }
  }
};
