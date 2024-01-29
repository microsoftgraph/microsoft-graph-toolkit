import { babel } from '@rollup/plugin-babel';
import json from '@rollup/plugin-json';
import nodeResolve from '@rollup/plugin-node-resolve';
import terser from '@rollup/plugin-terser';
import commonjs from '@rollup/plugin-commonjs';
import postcss from 'rollup-plugin-postcss';

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
  nodeResolve({ mainFields: ['module', 'jsnext'], extensions }),
  postcss(),
  terser({ keep_classnames: true, keep_fnames: true })
];

const es6Bundle = {
  input: ['src/index.ts'],
  output: {
    dir: 'dist/bundle',
    entryFileNames: 'mgt.js',
    format: 'esm',
    sourcemap: false
  },
  plugins: [
    ...commonPlugins,
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
    })
  ]
};

export default [es6Bundle];
