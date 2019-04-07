import resolve from 'rollup-plugin-node-resolve';
import babel from 'rollup-plugin-babel';
import typescript from "rollup-plugin-typescript";
import commonJS from "rollup-plugin-commonjs";
import postcss from "rollup-plugin-postcss";
import minify from 'rollup-plugin-babel-minify';

const extensions = [".js", ".ts"];

const commonPlugins = [
    commonJS(),
    resolve({ module: true, jsnext: true, extensions }),
    postcss()
    // minify({
    //     comments: false,
    //     sourceMap: false
    // })
    // terser({ keep_classnames: true, keep_fnames: true })
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
        typescript(),
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
            presets: ["@babel/env", "@babel/typescript"],
            plugins: [
                [
                    "@babel/plugin-proposal-decorators",
                    { decoratorsBeforeExport: true, legacy: false }
                ],
                "@babel/proposal-class-properties",
                "@babel/proposal-object-rest-spread"
            ],
            include: [
                "src/**/*",
                "node_modules/lit-element/**/*",
                "node_modules/lit-html/**/*"
            ],
        }),
        ...commonPlugins
    ]
}

export default [es6Bundle, es5Bundle];