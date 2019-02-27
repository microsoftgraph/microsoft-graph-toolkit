var path = require("path");
var webpack = require("webpack");

var PATHS = {
    entryPoint: path.resolve(__dirname, 'src/index.ts'),
    bundles: path.resolve(__dirname, 'dist/es5'),
}

var ForkTsCheckerWebpackPlugin = require('fork-ts-checker-webpack-plugin');

module.exports = {
    mode: "production",
    entry: {
        'msgraphproviders': [PATHS.entryPoint],
        'msgraphproviders.min': [PATHS.entryPoint]
    },
    output: {
        path: PATHS.bundles,
        filename: '[name].js',
        chunkFilename: '[name].chunk.js',
        libraryTarget: 'umd',
        library: 'msgraphtoolkit',
        umdNamedDefine: true
    },
    resolve: {
        extensions: ['.ts', '.tsx', '.js']
    },
    plugins: [
        new ForkTsCheckerWebpackPlugin()
    ],
    module: {
        rules: [{
            test: /\.(ts|tsx|js)?$/,
            use: {
                loader: 'ts-loader',
                options: {
                    transpileOnly: true,
                }
            }

        }]
    }
}