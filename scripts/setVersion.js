var child_process = require('child_process')
var path = require('path');
var fs = require('fs');
var project = require('../package.json');

const getPackageJsons = (startPath, results) => {
  startPath = startPath || 'packages';
  results = results || [];
  const filter = 'package.json';
  if (!fs.existsSync(startPath)) {
    console.log('no dir ', startPath);
    return;
  }

  var files = fs.readdirSync(startPath);
  for (var i = 0; i < files.length; i++) {
    var filename = path.join(startPath, files[i]);
    var stat = fs.lstatSync(filename);
    if (stat.isDirectory()) {
      getPackageJsons(filename, results); //recurse
    } else if (filename.indexOf(filter) >= 0) {
      results.push(filename);
    }
  }

  return results;
}

const updateMgtDependencyVersion = (packages, version) => {
  for (let package of packages) {
    const data = fs.readFileSync(package, 'utf8');

    var result = data.replace(/"(@microsoft\/mgt.*)": "(\*)"/g, `"$1": "${version}"`);
    result = result.replace(/"version": "(.*)"/g, `"version": "${version}"`);

    fs.writeFileSync(package, result, 'utf8');
  }
}

let version = project.version;

if (process.argv.length > 2) {
  switch (process.argv[2]) {
    case '-n':
    case '--next':
      // set version from git hash
      const shortSha = child_process.execSync('git rev-parse --short HEAD').toString().trim();
      version = `${version}-preview.${shortSha}`;
      break;
    case '-v':
    case '--version':
      // set version from argument
      if (process.argv.length > 3) {
        version = process.argv[3];
        break;
      }
    default:
      console.log('usage: node setVersion.js');
      console.log('usage: node setVersion.js --next');
      console.log('usage: node setVersion.js --version [version]');
      return;
  }
}

const packages = getPackageJsons();
updateMgtDependencyVersion(packages, version);
