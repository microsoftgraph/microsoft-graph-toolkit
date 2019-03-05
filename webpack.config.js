var path = require("path");
var ForkTsCheckerWebpackPlugin = require('fork-ts-checker-webpack-plugin');

var PATHS = {
    entryPoint: path.resolve(__dirname, 'es5/index.ts'),
    bundles: path.resolve(__dirname, 'dist/es5'),
}

module.exports = {
    mode: "production",
    entry: {
        'ms-graph-toolkit': [PATHS.entryPoint]
    },
    output: {
        path: PATHS.bundles,
        filename: '[name].js',
        chunkFilename: '[name].chunk.js',
        libraryTarget: 'umd',
        library: 'msGraphToolkit',
        umdNamedDefine: true
    },
    resolve: {
        extensions: ['.ts', '.js']
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
                    compilerOptions : {
                        target: "es5"
                    }
                }
            }

        }]
    }
}