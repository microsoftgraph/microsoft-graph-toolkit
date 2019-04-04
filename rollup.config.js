import resolve from "rollup-plugin-node-resolve";
import commonJS from "rollup-plugin-commonjs";
import postcss from "rollup-plugin-postcss";
import { terser } from "rollup-plugin-terser";
import typescript from "rollup-plugin-typescript";
import babel from "rollup-plugin-babel";

const extensions = [".js", ".ts"];

const src_root = "./src";
const bin_root = "./dist";

const es6_bundle_inputs = `${src_root}/index.ts`;
const es5_bundle_input = `${src_root}/index.es5.ts`;

const base_plugins = [
  commonJS(),
  resolve({ module: true, jsnext: true, extensions }),
  postcss({ inject: false }),
  terser({ keep_classnames: true, keep_fnames: true })
];

const babel_es5 = {
  extensions,
  presets: ["@babel/env", "@babel/typescript"],
  plugins: [
    "@babel/proposal-class-properties",
    "@babel/proposal-object-rest-spread",
    [
      "@babel/plugin-proposal-decorators",
      { decoratorsBeforeExport: true, legacy: false }
    ]
  ],
  include: [
    "src/**/*",
    "node_modules/lit-element/**/*"
  ],
  exclude: [
    "node_modules/@webcomponents/webcomponentsjs/custom-elements-es5-adapter.js"
  ]
};

const typescript_es6 = {
  target: "es2015",
  module: "esnext",
  moduleResolution: "node",
  lib: ["dom", "es2015"],
  allowSyntheticDefaultImports: true,
  experimentalDecorators: true,
  esModuleInterop: true
};

const typescript_es5 = {
  ...typescript_es6,
  target: "es5",
  module: "esnext"
};

export default [
  // ES6 Bundle
  {
    input: es6_bundle_inputs,
    plugins: [typescript(typescript_es6), ...base_plugins],
    output: {
      dir: `${bin_root}/es6/bundle`,
      entryFileNames: "mgt.js",
      format: "esm"
    }
  },
  // ES5 Bundle
  {
    input: es5_bundle_input,
    plugins: [babel(babel_es5), ...base_plugins],
    output: {
      dir: `${bin_root}/es5/bundle`,
      name: `index`,
      entryFileNames: "mgt.js",
      format: "iife"
    }
  }
];
