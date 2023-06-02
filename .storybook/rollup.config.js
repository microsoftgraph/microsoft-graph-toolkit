import resolve from 'rollup-plugin-node-resolve';
import babel from 'rollup-plugin-babel';
import postcss from 'rollup-plugin-postcss';
import { terser } from 'rollup-plugin-terser';

const extensions = ['.js'];

const getBabelConfig = isEs5 => {
  return {
    plugins: [
      ['@babel/plugin-proposal-decorators', isEs5 ? { legacy: true } : { decoratorsBeforeExport: true, legacy: false }],
      ['@babel/proposal-class-properties', { loose: isEs5 }],
      '@babel/proposal-object-rest-spread'
    ],
    include: [
      'packages/**/*',
      'node_modules/lit-element/**/*',
      'node_modules/lit-html/**/*',
      'node_modules/@microsoft/microsoft-graph-client/lib/es/**/*',
      'node_modules/msal/lib-es6/**/*'
    ]
  };
};

const es6Bundle = {
  input: ['./packages/mgt/dist/es6/index.js'],
  output: {
    dir: 'assets',
    entryFileNames: 'mgt.storybook.js',
    format: 'esm'
  },
  plugins: [
    babel({
      extensions,
      ...getBabelConfig(false)
    }),
    resolve({ module: true, jsnext: true, extensions }),
    postcss(),
    terser({ keep_classnames: true, keep_fnames: true })
  ]
};

const es6PreviewBundle = {
  input: ['./packages/mgt/dist/es6/preview-bundle.js'],
  output: {
    dir: 'assets',
    entryFileNames: 'mgt.preview.storybook.js',
    format: 'esm'
  },
  plugins: [
    babel({
      extensions,
      ...getBabelConfig(false)
    }),
    resolve({ module: true, jsnext: true, extensions }),
    postcss(),
    terser({ keep_classnames: true, keep_fnames: true })
  ]
};

export default [es6Bundle, es6PreviewBundle];
