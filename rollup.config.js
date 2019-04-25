import resolve from 'rollup-plugin-node-resolve';
import babel from 'rollup-plugin-babel';
import commonJS from "rollup-plugin-commonjs";
import postcss from "rollup-plugin-postcss";
import { terser } from "rollup-plugin-terser";
import json from 'rollup-plugin-json';

const extensions = [".js", ".ts"];

const commonPlugins = [
    json(),
    commonJS(),
    resolve({ module: true, jsnext: true, extensions }),
    postcss(),
    terser({ keep_classnames: true, keep_fnames: true })
];

const babelPlugins = [
    [
        "@babel/plugin-proposal-decorators",
        { decoratorsBeforeExport: true, legacy: false }
    ],
    "@babel/proposal-class-properties",
    "@babel/proposal-object-rest-spread"
];

const babelInclude = [
    "src/**/*",
    "node_modules/lit-element/**/*",
    "node_modules/lit-html/**/*",
    "node_modules/@microsoft/microsoft-graph-client/lib/es/**/*",
    "node_modules/msal/lib-es6/**/*"
];

const es6Bundle = {
	input: ['src/bundle/index.es6.ts'],
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
                    "@babel/preset-env", {
                        "targets": ">25%"
                    }
                ], "@babel/typescript"
            ],
            plugins: babelPlugins,
            include: babelInclude,
        }),
        ...commonPlugins
    ]
}

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
                    "@babel/preset-env", {
                        "targets": "last 2 versions"
                    }
                ], "@babel/typescript"
            ],
            plugins: babelPlugins,
            include: babelInclude,
        }),
        ...commonPlugins
    ]
}

export default [es6Bundle, es5Bundle];