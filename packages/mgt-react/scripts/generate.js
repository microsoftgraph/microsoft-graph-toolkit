var fs = require('fs-extra');

let wc = JSON.parse(fs.readFileSync(`${__dirname}/../temp/web-components.json`));

const primitives = new Set(['string', 'boolean', 'number', 'any']);
const mgtImports = new Set();

const tags = new Set([
  'mgt-person',
  'mgt-person-card',
  'mgt-agenda',
  'mgt-get',
  'mgt-login',
  'mgt-people-picker',
  'mgt-people',
  'mgt-tasks',
  'mgt-teams-channel-picker'
]);

let output = '';

const wrappers = [];

for (const tag of wc.tags) {
  if (!tags.has(tag.name)) {
    continue;
  }

  const className = tag.name
    .split('-')
    .slice(1)
    .map(t => t[0].toUpperCase() + t.substring(1))
    .join('');

  wrappers.push({
    tag: tag.name,
    propsType: className + 'Props',
    className: className
  });

  const props = {};

  for (const prop of tag.properties) {
    if (prop.type) {
      props[prop.name] = prop.type;

      let type = prop.type;
      if (type.endsWith('[]')) {
        type = type.substring(0, type.length - 2);
      } else if (type.startsWith('Array<')) {
        type = type.substring(6, type.length - 1);
      } else if (type === '*') {
        continue;
      }

      if (
        type.startsWith('MicrosoftGraph.') ||
        type.startsWith('MicrosoftGraphBeta.') ||
        type.startsWith('MgtElement.')
      ) {
        continue;
      }

      if (!primitives.has(type) && !mgtImports.has(type)) {
        console.log(type);
        mgtImports.add(type);
      }
    }
  }

  let propsType = '';

  for (const prop in props) {
    propsType += `\t${prop}?: ${props[prop]};\n`;
  }

  if (tag.events) {
    for (const event of tag.events) {
      propsType += `\t${event.name}?: (e: Event) => void;\n`;
    }
  }

  output += `\nexport type ${className}Props = {\n${propsType}}\n`;
}

for (const wrapper of wrappers) {
  output += `\nexport const ${wrapper.className} = wrapMgt<${wrapper.propsType}>('${wrapper.tag}');\n`;
}

output = `import { ${Array.from(mgtImports).join(',')} } from '@microsoft/mgt';
import * as MgtElement from '@microsoft/mgt-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import {wrapMgt} from '../Mgt';
${output}
`;

if (!fs.existsSync(`${__dirname}/../src/generated`)) {
  fs.mkdirSync(`${__dirname}/../src/generated`);
}

fs.writeFileSync(`${__dirname}/../src/generated/react.ts`, output);
