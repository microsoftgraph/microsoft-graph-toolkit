module.exports = {
  module: {
    rules: [
      {
        test: /\.scss$/,
        loader: 'postcss-loader',
        options: {
          postcssOptions: {
            plugins: ['postcss-import', 'tailwindcss', 'autoprefixer']
          }
        }
      }
    ]
  }
};
