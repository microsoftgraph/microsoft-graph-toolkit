import resolve from "rollup-plugin-node-resolve";
import commonJS from "rollup-plugin-commonjs";
import typescript from "rollup-plugin-typescript";
import postcss from "rollup-plugin-postcss";
import { terser } from "rollup-plugin-terser";

const src_root = "./src";
const bin_root = "./dist";

const bundle_inputs = `${src_root}/index.ts`;
const ui_inputs = `${src_root}/components/ui/ui.ts`;
const provider_inputs = [
  `${src_root}/components/providers/mgt-mock-provider.ts`,
  `${src_root}/components/providers/mgt-msal-provider.ts`,
  `${src_root}/components/providers/mgt-teams-provider.ts`,
  `${src_root}/components/providers/mgt-wam-provider.ts`
];

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
  target: "es5",
  module: "esnext",
  moduleResolution: "node",
  lib: ["dom", "es2015"],
  allowSyntheticDefaultImports: true,
  experimentalDecorators: true,
  esModuleInterop: true
};

const base_plugins = [
  commonJS(),
  resolve({ module: true, jsnext: true }),
  postcss({ inject: false }),
  terser({ keep_classnames: true, keep_fnames: true })
];

const es6_bundle_config = {
  input: bundle_inputs,
  plugins: [...base_plugins, typescript(typescript_es6)],
  output: {
    dir: `${bin_root}/es6/bundle`,
    entryFileNames: "[name].js",
    format: "esm"
  }
};

const es6_components_config = {
  input: [ui_inputs, ...provider_inputs],
  plugins: [...base_plugins, typescript(typescript_es6)],
  output: {
    dir: `${bin_root}/es6/components`,
    entryFileNames: "[name].js",
    format: "esm"
  }
};

const es5_bundle_config = {
  input: bundle_inputs,
  plugins: [...base_plugins, typescript(typescript_es5)],
  output: {
    dir: `${bin_root}/es5/bundle`,
    name: `index`,
    entryFileNames: "[name].js",
    format: "iife"
  }
};
const es5_components_config = {
  input: [ui_inputs, ...provider_inputs],
  plugins: [...base_plugins, typescript(typescript_es5)],
  output: {
    dir: `${bin_root}/es5/components`,
    entryFileNames: "[name].js",
    format: "iife"
  }
};

export default [
  es6_bundle_config,
  es6_components_config,
  es5_bundle_config,
  // es5_components_config
];
