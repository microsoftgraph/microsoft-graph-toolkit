const fs = require('fs');
const path = require('path');

function fixTitle(filePath, title) {
  const htmlDocumentPath = path.resolve(__dirname, filePath);
  const htmlDocument = fs.readFileSync(htmlDocumentPath, 'utf-8');
  const updatedHtmlDocument = htmlDocument.replace(/<title>.*<\/title>/, `<title>${title}</title>`);

  fs.writeFileSync(htmlDocumentPath, updatedHtmlDocument);
}

function addSeoTags(filePath, title) {
  const htmlDocumentPath = path.resolve(__dirname, filePath);
  const htmlDocument = fs.readFileSync(htmlDocumentPath, 'utf-8');
  const updatedHtmlDocument = htmlDocument.replace(
    /<\/title>/,
    `</title>
    <meta name="og:title" content="${title}"></meta>
    <meta name="og:type" content="website"></meta>
    <meta name="og:description" content="Use Microsoft Graph Toolkit to find authentication providers and reusable, framework-agnostic web components for accessing and working with Microsoft Graph."></meta>
    <meta name="description" content="Use Microsoft Graph Toolkit to find authentication providers and reusable, framework-agnostic web components for accessing and working with Microsoft Graph."></meta>
`
  );

  fs.writeFileSync(htmlDocumentPath, updatedHtmlDocument);
}

try {
  const args = process.argv.slice(2);
  const [title, distPath] = args;

  const storybookDistPath = `${distPath}storybook-static`;
  const indexPath = `${storybookDistPath}/index.html`;
  const iframePath = `${storybookDistPath}/iframe.html`;

  console.log(`Rewriting ${indexPath} document title to ${title}.`);
  fixTitle(indexPath, title);
  console.log(`Adding SeoTags to ${indexPath}.`);
  addSeoTags(indexPath, title);

  console.log(`Rewriting iframe.html document title to ${title}.`);
  fixTitle(iframePath, title);

  console.log('Title rewrite complete.');
} catch (error) {
  console.log('Title rewrite failed.');
  console.error(error);
  process.exit(1);
}
