import { makeDecorator } from '@storybook/addons';
import { EditorElement } from './editor';

const mgtScriptName = './mgt.storybook.js';

// function is used for dragging and moving
const setupEditorResize = (first, separator, last, dragComplete) => {
  var md; // remember mouse down info

  separator.onmousedown = onMouseDown;

  function onMouseDown(e) {
    // console.log('mouse down: ' + e.clientX);
    md = {
      e,
      offsetLeft: separator.offsetLeft,
      offsetTop: separator.offsetTop,
      firstWidth: first.offsetWidth,
      lastWidth: last.offsetWidth,
      firstHeight: first.offsetHeight,
      lastHeight: last.offsetHeight
    };
    document.onmousemove = onMouseMove;
    document.onmouseup = () => {
      if (typeof dragComplete === 'function') {
        dragComplete();
      }
      // console.log('mouse up');
      document.onmousemove = document.onmouseup = null;
    };
  }

  function onMouseMove(e) {
    // console.log('mouse move: ' + e.clientX);
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
  }
};

let scriptRegex = /<script\b[^>]*>([\s\S]*?)<\/script>/gm;
let styleRegex = /<style\b[^>]*>([\s\S]*?)<\/style>/gm;

export const withCodeEditor = makeDecorator({
  name: `withCodeEditor`,
  parameterName: 'myParameter',
  skipIfNoParametersOrOptions: false,
  wrapper: (getStory, context, { parameters }) => {
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

    let editor = new EditorElement();
    editor.files = {
      html: storyHtml,
      js: scriptCode,
      css: styleCode
    };

    editor.addEventListener('fileUpdated', () => {
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
                import {Providers, MockProvider} from "${mgtScriptName}";
                Providers.globalProvider = new MockProvider(true);
              </script>
              <style>
                html, body {
                  height: 100%;
                }
                ${css}
              </style>
            </head>
            <body>
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
      storyElementWrapper.innerHTML = '';
      storyElementWrapper.appendChild(storyElement);
    });

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
