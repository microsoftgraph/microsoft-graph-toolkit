const path = require('path');

module.exports = {
    entry: {
        auth: './lib/index.js'
    },
    output : {
        filename: 'auth.js',
        library: 'Auth',
        path: path.resolve(__dirname, 'dist')
    }
};