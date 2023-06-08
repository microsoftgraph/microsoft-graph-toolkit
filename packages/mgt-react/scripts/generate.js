var fs = require('fs-extra');

let wc = JSON.parse(fs.readFileSync(`${__dirname}/../temp/custom-elements.json`));

const primitives = new Set(['string', 'boolean', 'number', 'any', 'void', 'null', 'undefined']);

const gaTags = new Set([
  'person',
  'person-card',
  'agenda',
  'get',
  'login',
  'people-picker',
  'people',
  'tasks',
  'teams-channel-picker',
  'todo',
  'file',
  'file-list',
  'file-grid',
  'file-list-composite',
  'picker',
  'theme-toggle',
  'spinner'
]);
const outputFileName = 'react';

const previewTags = new Set(['search-box', 'search-results']);
const previewOutputFileName = 'react-preview';

const generateTags = (tags, fileName, importPreviewBarrel) => {
  const mgtComponentImports = new Set();
  const mgtElementImports = new Set();
  let output = '';

  const wrappers = [];

  const customTags = [];
  for (const module of wc.modules) {
    for (const d of module.declarations) {
      if (d.customElement && d.tagName && tags.has(d.tagName)) {
        customTags.push(d);
      }
    }
  }

  const removeGenericTypeDecoration = type => {
    if (type.endsWith('[]')) {
      return type.substring(0, type.length - 2);
    } else if (type.startsWith('Array<')) {
      return removeGenericTypeDecoration(type.substring(6, type.length - 1));
    } else if (type.startsWith('CustomEvent<')) {
      return removeGenericTypeDecoration(type.substring(12, type.length - 1));
    }
    return type;
  };

  const addTypeToImports = type => {
    if (type === '*') {
      return;
    }
    // make sure to remove any generic type decorations before trying to split for union types
    type = removeGenericTypeDecoration(type);

    // need to handle Generic like Command<T>
    if (type.indexOf('<') > 0 && type.indexOf('>') === type.length - 1) {
      const genericType = type.substring(0, type.indexOf('<'));
      const genericTypeArgs = type.substring(type.indexOf('<') + 1, type.length - 2);
      addTypeToImports(genericType);
      addTypeToImports(genericTypeArgs);
      return;
    }

    for (let t of type.split('|')) {
      t = removeGenericTypeDecoration(t.trim());
      if (t.startsWith('MicrosoftGraph.') || t.startsWith('MicrosoftGraphBeta.')) {
        return;
      }

      if (t.startsWith('MgtElement.') && !mgtElementImports.has(t)) {
        mgtElementImports.add(t.split('.')[1]);
      } else if (!primitives.has(t) && !mgtComponentImports.has(t)) {
        mgtComponentImports.add(t);
      }
    }
  };

  for (const tag of customTags.sort((a, b) => (a.tagName > b.tagName ? 1 : -1))) {
    const className = tag.tagName
      .split('-')
      .map(t => t[0].toUpperCase() + t.substring(1))
      .join('');

    wrappers.push({
      tag: tag.tagName,
      propsType: className + 'Props',
      className: className
    });

    const props = {};

    for (let i = 0; i < tag.members.length; ++i) {
      const prop = tag.members[i];
      let type = prop.type?.text;

      if (type && prop.kind === 'field' && prop.privacy === 'public' && !prop.static) {
        if (prop.name) {
          props[prop.name] = type;
        }

        if (type.includes('|')) {
          const types = type.split('|');
          for (const t of types) {
            tag.members.push({
              kind: 'field',
              privacy: 'public',
              type: { text: t.trim() }
            });
          }
          continue;
        }

        addTypeToImports(type);
      }
    }

    let propsType = '';

    for (const prop in props) {
      let type = props[prop];
      if (type.startsWith('MgtElement.')) {
        type = type.split('.')[1];
      }
      propsType += `\t${prop}?: ${type};\n`;
    }

    if (tag.events) {
      for (const event of tag.events) {
        if (event.type && event.type.text) {
          // remove MgtElement. prefix as this it only used to ensure it's imported from the correct package
          propsType += `\t${event.name}?: (e: ${event.type.text.replace(
            'CustomEvent<MgtElement.',
            'CustomEvent<'
          )}) => void;\n`;
          // also ensure that the necessary import is added to either mgt-element or mgt-component imports
          addTypeToImports(event.type.text);
        } else {
          propsType += `\t${event.name}?: (e: Event) => void;\n`;
        }
      }
    }

    output += `\nexport type ${className}Props = {\n${propsType}}\n`;
  }

  for (const wrapper of wrappers) {
    output += `\nexport const ${wrapper.className} = wrapMgt<${wrapper.propsType}>('${wrapper.tag}');\n`;
  }

  output = `${importPreviewBarrel ? `import '@microsoft/mgt-components/dist/es6/components/preview';` : ''}
import { ${Array.from(mgtComponentImports).join(',')} } from '@microsoft/mgt-components';
import { ${Array.from(mgtElementImports).join(',')} } from '@microsoft/mgt-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import {wrapMgt} from '../Mgt';
${output}
`;

  if (!fs.existsSync(`${__dirname}/../src/generated`)) {
    fs.mkdirSync(`${__dirname}/../src/generated`);
  }

  fs.writeFileSync(`${__dirname}/../src/generated/${fileName}.ts`, output);
};

generateTags(gaTags, outputFileName, false);
generateTags(previewTags, previewOutputFileName, true);
