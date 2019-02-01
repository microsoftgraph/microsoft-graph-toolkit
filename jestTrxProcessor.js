var builder = require('jest-trx-results-processor');

var processor = builder({
    outputFile: 'jestTestresults.trx' 
});

module.exports = processor;