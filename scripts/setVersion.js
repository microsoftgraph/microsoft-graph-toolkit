var path = require('path'),
  fs = require('fs');

function getPackageJsons(startPath, results) {
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

function updateMgtDependencyVersion(packages, version) {
  for (let package of packages) {
    const data = fs.readFileSync(package, 'utf8');

    var result = data.replace(/"(@microsoft\/mgt.*)": "(\*)"/g, `"$1": "${version}"`);
    result = result.replace(/"version": "(.*)"/g, `"version": "${version}"`);

    fs.writeFileSync(package, result, 'utf8');
  }
}

if (process.argv.length < 3) {
  console.log('usage: node setVersion.js VERSION');
  return;
}

const version = process.argv[2];

packages = getPackageJsons();
updateMgtDependencyVersion(packages, version);
