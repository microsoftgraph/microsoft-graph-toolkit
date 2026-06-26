import { addons, makeDecorator } from '@storybook/preview-api';

import { ProviderState } from '../../../packages/mgt-element/dist/es6/providers/IProvider';
import { EditorElement } from './editor';
import { CLIENTID, SETPROVIDER_EVENT, AUTH_PAGE } from '../../env';
import { beautifyContent } from '../../utils/beautifyContent';
import { isValidManifestUrl } from '../../utils/isValidManifestUrl';

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

let reactRegex = /<react\b[^>]*>([\s\S]*?)<\/react>/gm;
let scriptRegex = /<script\b[^>]*>([\s\S]*?)<\/script>/gm;
let styleRegex = /<style\b[^>]*>([\s\S]*?)<\/style>/gm;

export const withCodeEditor = makeDecorator({
  name: `withCodeEditor`,
  parameterName: 'myParameter',
  skipIfNoParametersOrOptions: false,
  wrapper: (getStory, context, { options }) => {
    const forOptions = options ? options.disableThemeToggle : false;
    const title =
      ['Custom CSS Properties', 'Theme'].includes(context.name) || context.title.toLowerCase().includes('templating');
    const forContext = context && title;
    const disableThemeToggle = forOptions || forContext;
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

    let reactMatches = reactRegex.exec(storyHtml);
    let reactCode = reactMatches && reactMatches.length > 1 ? reactMatches[1].trim() : '';

    let styleMatches = styleRegex.exec(storyHtml);
    let styleCode = styleMatches && styleMatches.length > 1 ? styleMatches[1].trim() : '';

    storyHtml = storyHtml
      ?.replace(styleRegex, '')
      ?.replace(reactRegex, '')
      ?.replace(scriptRegex, '')
      ?.replace(/\n?<!---->\n?/g, '')
      ?.trim();

    const fileTypes = reactCode ? ['react', 'css'] : ['html', 'js', 'css'];

    let editor = new EditorElement(fileTypes);

    const isEditorEnabled = () => {
      return !context.parameters.docs?.editor?.hidden;
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
          console.warn(`🦒: Can't get content from '${url}'`);
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

    const isValid = isValidManifestUrl;

    if (context.name === 'Editor') {
      // If the editor is not iframed (Docs, GE, etc.)
      if (isNotIframed()) {
        var urlParams = new URLSearchParams(window.top.location.search);
        var manifestUrl = urlParams.get('manifest');

        if (isValid(manifestUrl)) {
          getContent(manifestUrl, true).then(manifest => {
            const contentUrls = [manifest[0].preview.html, manifest[0].preview.js, manifest[0].preview.css];
            if (contentUrls.some(u => u && !isValid(u))) {
              console.warn(`🦒: Manifest contains untrusted URLs`);
              return;
            }
            Promise.all([
              getContent(manifest[0].preview.html),
              getContent(manifest[0].preview.js),
              getContent(manifest[0].preview.css)
            ]).then(values => {
              //editor.autoFormat = false;
              editor.files = {
                html: beautifyContent('html', values[0]),
                js: beautifyContent('js', values[1]),
                css: beautifyContent('css', values[2])
              };
            });
          });
        }
      }
    }

    const themeToggleCss = disableThemeToggle
      ? ''
      : `
      body {
        background-color: var(--fill-color);
        color: var(--neutral-foreground-rest);
        font-family: var(--body-font);
        padding: 0 12px;
      }
      header {
        display: flex;
        flex-direction: row;
        justify-content: flex-end;
        padding: 0 0 12px 0;
      }
`;
    const themeToggle = disableThemeToggle
      ? ''
      : `
      <header>
        <mgt-theme-toggle mode="light"></mgt-theme-toggle>
      </header>
`;

    let providerInitCode = `
      import {Providers, MockProvider} from "${mgtScriptName}";
      Providers.globalProvider = new MockProvider(true);
    `;

    const channel = addons.getChannel();
    channel.on(SETPROVIDER_EVENT, params => {
      if (params.state === ProviderState.SignedIn && params.name === 'MgtMockProvider') {
        providerInitCode = `
          import { Providers, MockProvider } from "${mgtScriptName}";
          Providers.globalProvider = new MockProvider(true);
        `;
      } else if (params.state === ProviderState.SignedIn && params.name === 'MgtMsal2Provider') {
        providerInitCode = `
          import { Providers, Msal2Provider, LoginType } from "${mgtScriptName}";
          Providers.globalProvider = new Msal2Provider({
            clientId: "${CLIENTID}",
            loginType: LoginType.Popup,
            redirectUri: "${window.location.origin}/${AUTH_PAGE}"
          });`;
      }
    });

    const getStoryTitle = context => {
      const storyTitle = `${context?.title} - ${context?.story}`;
      return storyTitle;
    };

    const loadEditorContent = () => {
      const storyElement = document.createElement('iframe');

      // Security: sandbox the iframe to restrict capabilities.
      // allow-same-origin is required for ES module loading; exfiltration is
      // blocked by the CSP meta tag injected below.
      storyElement.setAttribute(
        'sandbox',
        'allow-scripts allow-same-origin allow-popups allow-popups-to-escape-sandbox allow-forms'
      );

      storyElement.addEventListener(
        'load',
        () => {
          let doc = storyElement.contentDocument;

          let { html, css, js } = editor.files;
          js = js.replace(
            /import \{([^\}]+)\}\s+from\s+['"]@microsoft\/mgt\x2d.*['"];/gm,
            `import {$1} from '${mgtScriptName}';`
          );

          const docContent = `
            <html>
              <head>
                <meta http-equiv="Content-Security-Policy"
                  content="default-src 'self';
                    script-src 'self' 'unsafe-inline';
                    style-src 'self' 'unsafe-inline';
                    connect-src https://graph.microsoft.com https://graph.microsoft.us https://dod-graph.microsoft.us https://graph.microsoft.de https://microsoftgraph.chinacloudapi.cn https://canary.graph.microsoft.com https://login.microsoftonline.com https://cdn.graph.office.net https://graph.office.net 'self';
                    img-src 'self' data: blob: https://*.microsoft.com https://*.microsoftonline.com https://*.sharepoint.com https://*.office.com https://*.office365.com https://*.windows.net;
                    font-src 'self' https://static2.sharepointonline.com;
                    frame-src https://login.microsoftonline.com 'self';
                    form-action 'none';
                    object-src 'none';
                    base-uri 'self';"
                />
                <script type="module" src="${mgtScriptName}"></script>
                <script type="module">
                  import { registerMgtComponents } from "${mgtScriptName}";
                  registerMgtComponents();
                </script>
                <script type="module">
                  ${providerInitCode}
                </script>
                <style>
                  ${themeToggleCss}
                  ${css}
                </style>
              </head>
              <body>
                ${themeToggle}
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
        },
        { once: true }
      );

      storyElement.className = 'story-mgt-preview';
      storyElement.setAttribute('title', 'preview');
      storyElementWrapper.innerHTML = '';
      storyElementWrapper.appendChild(storyElement);
    };

    editor.addEventListener('fileUpdated', loadEditorContent);

    const separator = document.createElement('div');

    setupEditorResize(storyElementWrapper, separator, editor, () => editor.layout());

    root.className = 'story-mgt-root';

    storyElementWrapper.className = isEditorEnabled() ? 'story-mgt-preview-wrapper' : 'story-mgt-preview-wrapper-full';
    separator.className = 'story-mgt-separator';
    editor.className = 'story-mgt-editor';

    root.appendChild(storyElementWrapper);
    root.appendChild(separator);

    if (isEditorEnabled()) {
      root.appendChild(editor);
    }

    window.addEventListener('resize', () => {
      storyElementWrapper.style.height = null;
      storyElementWrapper.style.width = null;
      editor.style.height = null;
      editor.style.width = null;
    });

    editor.files = {
      html: beautifyContent('html', storyHtml),
      react: beautifyContent('js', reactCode),
      js: beautifyContent('js', scriptCode),
      css: beautifyContent('css', styleCode)
    };

    editor.title = getStoryTitle(context);

    return root;
  }
});
