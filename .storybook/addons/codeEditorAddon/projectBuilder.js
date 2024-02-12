import sdk from '@stackblitz/sdk';
import { beautifyContent } from '../../utils/beautifyContent';
import { getCleanVersionInfo } from '../../versionInfo';

const TEMPLATE_PATH = [window.location.protocol, '//', window.location.host, window.location.pathname]
  .join('')
  .replace('iframe.html', '');
const REACT_TEMPLATE_PATH = TEMPLATE_PATH + 'stackblitz/react/';
const REACT_TEMPLATE_FILES = [
  { name: 'package.json', type: 'json' },
  { name: 'index.html', type: 'html' },
  { name: 'tsconfig.json', type: 'json' },
  { name: 'tsconfig.node.json', type: 'json' },
  { name: 'vite.config.ts', type: 'js' },
  { name: 'src/vite-env.d.ts', type: 'js' },
  { name: 'src/main.css', type: 'css' },
  { name: 'src/main.tsx', type: 'js' }
];
const HTML_TEMPLATE_PATH = TEMPLATE_PATH + 'stackblitz/html/';
const HTML_TEMPLATE_FILES = [
  { name: 'index.html', type: 'html' },
  { name: 'main.css', type: 'css' },
  { name: 'app.js', type: 'js' },
  { name: 'main.js', type: 'js' },
  { name: 'package.json', type: 'json' },
  { name: 'style.css', type: 'css' }
];

export const generateProject = async (title, files) => {
  files.react && files.react !== '\n' ? await openReactProject(title, files) : await openHtmlProject(title, files);
};

let openReactProject = async (title, files) => {
  const snippets = [
    { name: 'src/App.tsx', type: 'js', content: files.react },
    { name: 'src/App.css', type: 'css', content: files.css }
  ];
  const stackblitzFiles = await buildFiles(REACT_TEMPLATE_PATH, REACT_TEMPLATE_FILES, snippets);
  sdk.openProject(
    {
      files: stackblitzFiles,
      template: 'node',
      title: title
    },
    {
      openFile: 'src/App.tsx',
      newWindow: true
    }
  );
};

let openHtmlProject = async (title, files) => {
  const snippets = [
    { name: 'index.html', type: 'html', content: files.html },
    { name: 'style.css', type: 'css', content: files.css },
    { name: 'main.js', type: 'js', content: files.js }
  ];
  const stackblitzFiles = await buildFiles(HTML_TEMPLATE_PATH, HTML_TEMPLATE_FILES, snippets);
  sdk.openProject(
    {
      files: stackblitzFiles,
      template: 'node',
      title: title
    },
    {
      openFile: 'index.html',
      newWindow: true
    }
  );
};

let buildFiles = async (templatePath, files, snippets) => {
  const stackblitzFiles = {};
  await Promise.all(
    files.map(async file => {
      let fileContent = await loadFile(templatePath + file.name);
      fileContent = fileContent.replace(/<mgt-version><\/mgt-version>/g, getCleanVersionInfo());
      stackblitzFiles[file.name] = beautifyContent(file.type, fileContent);
    })
  );

  snippets.map(file => {
    let fileContent = '';
    if (file.content) {
      if (stackblitzFiles[file.name] && stackblitzFiles[file.name] !== '\n') {
        fileContent = stackblitzFiles[file.name].replace(/<mgt-component><\/mgt-component>/g, file.content);
      } else {
        fileContent = file.content;
      }
    }

    fileContent = beautifyContent(file.type, fileContent);
    stackblitzFiles[file.name] = fileContent;
  });

  return stackblitzFiles;
};

let loadFile = async fileUrl => {
  let content = '';
  if (fileUrl) {
    let response = await fetch(fileUrl);

    if (response.ok) {
      content = await response.text();
    } else {
      console.warn(`🦒: Can't get snippet from '${fileUrl}'`);
    }
  }
  return content;
};
