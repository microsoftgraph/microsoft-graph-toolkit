import resolve from 'rollup-plugin-node-resolve';
import babel from 'rollup-plugin-babel';
import commonjs from 'rollup-plugin-commonjs';
import postcss from 'rollup-plugin-postcss';
import { terser } from 'rollup-plugin-terser';
import json from 'rollup-plugin-json';

const extensions = ['.js', '.ts'];

const getBabelConfig = isEs5 => {
  return {
    plugins: [
      ['@babel/plugin-proposal-decorators', isEs5 ? { legacy: true } : { decoratorsBeforeExport: true, legacy: false }],
      ['@babel/proposal-class-properties', { loose: isEs5 }],
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
    format: 'iife',
    name: 'mgt',
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
      ...getBabelConfig(false)
    }),
    ...commonPlugins
  ]
};
const es6PreviewBundle = {
  input: ['src/preview.ts'],
  output: {
    dir: 'dist/bundle',
    entryFileNames: 'mgt.preview.es6.js',
    format: 'iife',
    name: 'mgt',
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
      ...getBabelConfig(false)
    }),
    ...commonPlugins
  ]
};

const es5Bundle = {
  input: ['src/bundle/index.es5.ts'],
  output: {
    dir: 'dist/bundle',
    entryFileNames: 'mgt.es5.js',
    format: 'iife',
    name: 'mgt',
    sourcemap: false
  },
  plugins: [
    babel({
      extensions,
      presets: [
        [
          '@babel/preset-env',
          {
            targets: 'last 2 versions',
            useBuiltIns: 'entry',
            corejs: 3
          }
        ],
        '@babel/typescript'
      ],
      ...getBabelConfig(true)
    }),
    ...commonPlugins
  ]
};

const es5PreviewBundle = {
  input: ['src/bundle/preview.es5.ts'],
  output: {
    dir: 'dist/bundle',
    entryFileNames: 'mgt.preview.es5.js',
    format: 'iife',
    name: 'mgt',
    sourcemap: false
  },
  plugins: [
    babel({
      extensions,
      presets: [
        [
          '@babel/preset-env',
          {
            targets: 'last 2 versions',
            useBuiltIns: 'entry',
            corejs: 3
          }
        ],
        '@babel/typescript'
      ],
      ...getBabelConfig(true)
    }),
    ...commonPlugins
  ]
};

const cjsBundle = {
  input: ['src/bundle/index.es5.ts'],
  output: {
    dir: 'dist/commonjs',
    entryFileNames: 'index.js',
    format: 'cjs',
    sourcemap: true
  },
  plugins: [
    babel({
      extensions,
      presets: [
        [
          '@babel/preset-env',
          {
            targets: 'last 2 versions'
          }
        ],
        '@babel/typescript'
      ],
      ...getBabelConfig(true)
    }),
    ...commonPlugins
  ]
};

const cjsPreviewBundle = {
  input: ['src/bundle/preview.es5.ts'],
  output: {
    dir: 'dist/commonjs',
    entryFileNames: 'preview.js',
    format: 'cjs',
    sourcemap: true
  },
  plugins: [
    babel({
      extensions,
      presets: [
        [
          '@babel/preset-env',
          {
            targets: 'last 2 versions'
          }
        ],
        '@babel/typescript'
      ],
      ...getBabelConfig(true)
    }),
    ...commonPlugins
  ]
};
export default [es6Bundle, es6PreviewBundle, es5Bundle, es5PreviewBundle, cjsBundle, cjsPreviewBundle];
// export default [es6Bundle, es6PreviewBundle, es5Bundle, cjsBundle];
