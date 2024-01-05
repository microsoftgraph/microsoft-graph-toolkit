import sdk from '@stackblitz/sdk';
import * as beautify from 'js-beautify';

export const generateProject = async (title, files) => {
  files.react ? await openReactProject(title, files) : await openHtmlProject(title, files);
};

const REACT_TEMPLATE_PATH = '/stackblitz/react/';
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
const HTML_TEMPLATE_PATH = '/stackblitz/html/';
const HTML_TEMPLATE_FILES = [
  { name: 'index.html', type: 'html' },
  { name: 'main.css', type: 'css' },
  { name: 'main.js', type: 'js' },
  { name: 'package.json', type: 'json' },
  { name: 'style.css', type: 'css' }
];

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
      const fileContent = await loadFile(templatePath + file.name);
      stackblitzFiles[file.name] = beautifyContent(file, fileContent);
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

    fileContent = beautifyContent(file, fileContent);
    stackblitzFiles[file.name] = fileContent;
  });

  return stackblitzFiles;
};

let beautifyContent = (file, content) => {
  const options = {
    indent_size: '2',
    indent_char: ' ',
    max_preserve_newlines: '0',
    preserve_newlines: true,
    keep_array_indentation: true,
    break_chained_methods: false,
    indent_scripts: 'separate',
    brace_style: 'collapse,preserve-inline',
    space_before_conditional: false,
    unescape_strings: false,
    jslint_happy: false,
    end_with_newline: true,
    wrap_line_length: '120',
    indent_inner_html: true,
    comma_first: false,
    e4x: true,
    indent_empty_lines: false
  };

  let beautifiedContent = content;
  switch (file.type) {
    case 'html':
      beautifiedContent = beautify.default.html(content, options);
      break;
    case 'js':
    case 'json':
      beautifiedContent = beautify.default.js(content, options);
      break;
    case 'css':
      beautifiedContent = beautify.default.css(content, options);
      break;
    default:
      break;
  }
  return beautifiedContent;
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
