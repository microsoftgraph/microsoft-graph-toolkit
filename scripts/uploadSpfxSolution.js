// uploads the mgt-spfx solution to a GitHub release
// this script is called by a github workflow to automate
// uploading the solution during a release of mgt

const { Octokit } = require("@octokit/rest");
const fs = require('fs');

if (process.argv.length < 4) {
  console.log('usage: uploadSpfxSolution.js <github_token> <package_version>');

  return;
}

const auth = process.argv[2];
const version = process.argv[3];
const repo = 'microsoft-graph-toolkit';
const owner = 'microsoftgraph';
const assetPath = 'packages/mgt-spfx/sharepoint/solution/mgt-spfx.sppkg';

const octokit = new Octokit({ auth });

(async () => {

  const releases = await octokit.rest.repos.listReleases({
    owner,
    repo
  })

  const filteredReleases = releases.data.filter(r => r.tag_name.includes(version));
  let release;

  if (!filteredReleases.length) {
    console.log(`No release found for ${version} - creating a new draft release`);
    release = (await octokit.rest.repos.createRelease({
      owner,
      repo,
      tag_name: 'v' + version,
      name: 'v' + version,
      draft: true
    })).data;
  } else {
    console.log(`found existing release for ${version}`);
    release = filteredReleases[0];
  }

  const file = fs.readFileSync(assetPath);
  const majorVersion = version.substring(0, version.indexOf("."));
  const filename = `mgt-spfx-v${majorVersion}.sppkg`;

  if (release.assets && release.assets.length) {
    const asset = release.assets.filter(a => a.name === name)[0];
    if (asset) {
      console.log(`found existing asset for release ${version} - deleting`)
      await octokit.rest.repos.deleteReleaseAsset({
        owner,
        repo,
        asset_id: asset.id
      });
    }
  }

  console.log(`attaching ${assetPath} as ${filename} to release`)
  await octokit.rest.repos.uploadReleaseAsset({
    owner,
    repo,
    release_id: release.id,
    data: file,
    name: filename
  });

  console.log('done')
})();