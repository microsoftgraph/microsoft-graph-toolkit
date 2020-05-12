const fs = require('fs-extra');

const temp = `${__dirname}/../temp`;
const dist = `${__dirname}/../dist`;

fs.removeSync(temp);
fs.removeSync(dist);
