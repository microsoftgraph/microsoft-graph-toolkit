import resolve from 'rollup-plugin-node-resolve';
import babel from 'rollup-plugin-babel';
import commonjs from 'rollup-plugin-commonjs';
import postcss from 'rollup-plugin-postcss';
import { terser } from 'rollup-plugin-terser';
import json from 'rollup-plugin-json';

const extensions = ['.js', '.ts'];

const getBabelConfig = () => {
  return {
    plugins: [
      ['@babel/plugin-proposal-decorators', { decoratorsBeforeExport: true, legacy: false }],
      ['@babel/proposal-class-properties', { loose: false }],
      '@babel/proposal-object-rest-spread'
    ],
    include: [
      'src/**/*',
      'node_modules/lit-element/**/*',
      'node_modules/lit-html/**/*',
      'node_modules/@microsoft/microsoft-graph-client/lib/es/**/*',
      'node_modules/msal/lib-es6/**/*'
    ]
  };
};

const commonPlugins = [
  json(),
  commonjs(),
  resolve({ module: true, jsnext: true, extensions }),
  postcss(),
  terser({ keep_classnames: true, keep_fnames: true })
];

const es6Bundle = {
  input: ['src/index.ts'],
  output: {
    dir: 'dist/bundle',
    entryFileNames: 'mgt.es6.js',
    format: 'esm',
    sourcemap: false
  },
  plugins: [
    babel({
      extensions,
      presets: [
        [
          '@babel/preset-env',
          {
            targets: '>25%'
          }
        ],
        '@babel/typescript'
      ],
      ...getBabelConfig()
    }),
    ...commonPlugins
  ]
};

export default [es6Bundle];
