import { makeDecorator } from '@storybook/addons';
import { html } from 'lit-element';
import { EditorElement } from './editor';

const mgtScriptName = './mgt.storybook.js';

// function is used for dragging and moving
const setupEditorResize = (first, separator, last, dragComplete) => {
  var md; // remember mouse down info

  separator.addEventListener('mousedown', e => {
    md = {
      e,
      offsetLeft: separator.offsetLeft,
      offsetTop: separator.offsetTop,
      firstWidth: first.offsetWidth,
      lastWidth: last.offsetWidth,
      firstHeight: first.offsetHeight,
      lastHeight: last.offsetHeight
    };

    first.style.pointerEvents = 'none';
    last.style.pointerEvents = 'none';

    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
  });

  const onMouseUp = () => {
    if (typeof dragComplete === 'function') {
      dragComplete();
    }

    first.style.pointerEvents = '';
    last.style.pointerEvents = '';

    document.removeEventListener('mousemove', onMouseMove);
    document.removeEventListener('mouseup', onMouseUp);
  };

  const onMouseMove = e => {
    var delta = { x: e.clientX - md.e.x, y: e.clientY - md.e.y };

    if (window.innerWidth > 800) {
      // Horizontal
      // prevent negative-sized elements
      delta.x = Math.min(Math.max(delta.x, -md.firstWidth + 200), md.lastWidth - 200);

      first.style.width = md.firstWidth + delta.x - 0.5 + 'px';
      last.style.width = md.lastWidth - delta.x - 0.5 + 'px';
    } else {
      // Vertical
      // prevent negative-sized elements
      delta.y = Math.min(Math.max(delta.y, -md.firstHeight + 150), md.lastHeight - 150);

      first.style.height = md.firstHeight + delta.y - 0.5 + 'px';
      last.style.height = md.lastHeight - delta.y - 0.5 + 'px';
    }
  };
};

let scriptRegex = /<script\b[^>]*>([\s\S]*?)<\/script>/gm;
let styleRegex = /<style\b[^>]*>([\s\S]*?)<\/style>/gm;

export const withCodeEditor = makeDecorator({
  name: `withCodeEditor`,
  parameterName: 'myParameter',
  skipIfNoParametersOrOptions: false,
  wrapper: (getStory, context, { parameters }) => {
    const getStoryHtml = html => {
      var urlParams = new URLSearchParams(window.location.search);
      var paramsToReplace = [];
      paramsToReplace.push({ key: '{theme}', value: urlParams.get('theme') });
      paramsToReplace.push({ key: '{entityId}', value: urlParams.get('entityId') });
      paramsToReplace.push({ key: '{resource}', value: atob(urlParams.get('resource')) });
      paramsToReplace.push({ key: '{version}', value: urlParams.get('version') });

      paramsToReplace.forEach(param => {
        html = html.replaceAll(param.key, param.value);
      });

      return html;
    };

    const getDynamicStory = (type, story, storyHtml, keys) => {
      if (type === 'graph-explorer--get-single') {
        let template = [];
        let dynamicTemplate = `<tr>
                      <td><pre>{{ key }}</pre></td>
                      <td >{{ this.{{ key }} }}</td>
                    </tr>\n`;

        //story.render();

        if (keys && keys.find(key => key === 'displayName')) {
          template.push(dynamicTemplate.replaceAll('{{ key }}', 'displayName'));
        }

        if (keys && keys.find(key => key === 'id')) {
          template.push(dynamicTemplate.replaceAll('{{ key }}', 'id'));
        }

        keys.map(key => {
          if (!key.startsWith('@')) {
            template.push(dynamicTemplate.replaceAll('{{ key }}', key));
          }
        });

        storyHtml = storyHtml.replace('{dynamicTemplate}', template.slice(0, 3).join(''));
      }

      return storyHtml;
    };

    let story = getStory(context);

    let storyHtml;
    const root = document.createElement('div');
    let storyElementWrapper = document.createElement('div');

    if (story.strings) {
      storyHtml = story.strings[0];
    } else {
      storyHtml = story.innerHTML;
    }

    let scriptMatches = scriptRegex.exec(storyHtml);
    let scriptCode = scriptMatches && scriptMatches.length > 1 ? scriptMatches[1].trim() : '';

    let styleMatches = styleRegex.exec(storyHtml);
    let styleCode = styleMatches && styleMatches.length > 1 ? styleMatches[1].trim() : '';

    storyHtml = storyHtml
      .replace(styleRegex, '')
      .replace(scriptRegex, '')
      .replace(/\n?<!---->\n?/g, '')
      .trim();

    storyHtml = getStoryHtml(storyHtml);

    var urlParams = new URLSearchParams(window.location.search);

    storyHtml = getDynamicStory(
      context.id,
      story,
      storyHtml,
      urlParams.get('keys') ? JSON.parse(atob(urlParams.get('keys'))) : null
    );

    let editor = new EditorElement();
    editor.files = {
      html: storyHtml,
      js: scriptCode,
      css: styleCode
    };

    const getContent = async (url, json) => {
      let content = '';

      if (url) {
        let response = await fetch(url);

        if (response.ok) {
          if (json) {
            content = await response.json();
          } else {
            content = await response.text();
          }
        } else {
          console.warn(`Can't get content from '${url}'`);
        }
      }

      return content;
    };

    const isNotIframed = () => {
      try {
        return window.top.location.href != null || window.top.location.href != undefined;
      } catch (err) {
        return false;
      }
    };

    const isValid = manifestUrl => {
      return manifestUrl && manifestUrl.startsWith('https://raw.githubusercontent.com/pnp/mgt-samples/main/');
    };

    if (context.name === 'Editor') {
      // If the editor is not iframed (Docs, GE, etc.)
      if (isNotIframed()) {
        var urlParams = new URLSearchParams(window.top.location.search);
        var manifestUrl = urlParams.get('manifest');

        if (isValid(manifestUrl)) {
          getContent(manifestUrl, true).then(manifest => {
            Promise.all([
              getContent(manifest[0].preview.html),
              getContent(manifest[0].preview.js),
              getContent(manifest[0].preview.css)
            ]).then(values => {
              //editor.autoFormat = false;
              editor.files = {
                html: values[0],
                js: values[1],
                css: values[2]
              };
            });
          });
        }
      }
    }

    const loadEditorContent = () => {
      const parsedHash = new URLSearchParams(window.location.hash.substring(1));
      let requiresLogin = false;

      let providerInitCode = `
        import {Providers, MockProvider} from "${mgtScriptName}";
        Providers.globalProvider = new MockProvider(true);
      `;

      if (parsedHash.get('client_id') && parsedHash.get('login_hint')) {
        providerInitCode = `
          import {Providers, Msal2Provider, LoginType} from "${mgtScriptName}";
          
          Providers.globalProvider = new Msal2Provider({
            clientId: '${parsedHash.get('client_id')}',
            loginHint: '${parsedHash.get('login_hint')}',
            loginType: LoginType.Popup
          });
        `;
      } else if (parsedHash.get('client_id')) {
        providerInitCode = `
          import {Providers, Msal2Provider, LoginType} from "${mgtScriptName}";
          
          Providers.globalProvider = new Msal2Provider({
            clientId: '${parsedHash.get('client_id')}',
            loginType: LoginType.Popup
          });
        `;

        requiresLogin = true;
      }

      const storyElement = document.createElement('iframe');

      storyElement.addEventListener('load', () => {
        let doc = storyElement.contentDocument;

        let { html, css, js } = editor.files;
        js = js.replace(
          /import \{([^\}]+)\}\s+from\s+['"]@microsoft\/mgt['"];/gm,
          `import {$1} from '${mgtScriptName}';`
        );

        const docContent = `
          <html>
            <head>
              <script type="module" src="${mgtScriptName}"></script>
              <script type="module">
                ${providerInitCode}
              </script>
              <style>
                html, body {
                  height: 100%;
                }
                ${css}
              </style>
            </head>
            <body>
              ${requiresLogin ? '<mgt-login></mgt-login>' : ''}
              ${html}
              <script type="module">
                ${js}
              </script>
            </body>
          </html>
        `;

        doc.open();
        doc.write(docContent);
        doc.close();
      });

      storyElement.className = 'story-mgt-preview';
      storyElement.setAttribute('title', 'preview');
      storyElementWrapper.innerHTML = '';
      storyElementWrapper.appendChild(storyElement);
    };

    editor.addEventListener('fileUpdated', loadEditorContent);
    window.addEventListener('hashchange', loadEditorContent);

    const separator = document.createElement('div');

    setupEditorResize(storyElementWrapper, separator, editor, () => editor.layout());

    root.className = 'story-mgt-root';
    storyElementWrapper.className = 'story-mgt-preview-wrapper';
    separator.className = 'story-mgt-separator';
    editor.className = 'story-mgt-editor';

    root.appendChild(storyElementWrapper);
    root.appendChild(separator);
    root.appendChild(editor);

    window.addEventListener('resize', () => {
      storyElementWrapper.style.height = null;
      storyElementWrapper.style.width = null;
      editor.style.height = null;
      editor.style.width = null;
    });

    return root;
  }
});
