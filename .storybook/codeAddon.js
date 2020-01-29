import addons, { makeDecorator } from '@storybook/addons';
import './components/editor';
import { EditorElement } from './components/editor';

// function is used for dragging and moving
const setupDrag = (first, separator, last, dragComplete) => {
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
      delta.x = Math.min(Math.max(delta.x, -md.firstWidth), md.lastWidth);

      separator.style.left = md.offsetLeft + delta.x + 'px';
      first.style.width = md.firstWidth + delta.x + 'px';
      last.style.width = md.lastWidth - delta.x + 'px';
    } else {
      // Vertical
      // prevent negative-sized elements
      delta.y = Math.min(Math.max(delta.y, -md.firstHeight), md.lastHeight);

      separator.style.top = md.offsetTop + delta.y + 'px';
      first.style.height = md.firstHeight + delta.y + 'px';
      last.style.height = md.lastHeight - delta.y + 'px';
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
    let storyHtml = story.innerHTML;

    let scriptMatches = scriptRegex.exec(storyHtml);
    let scriptCode = scriptMatches && scriptMatches.length > 1 ? scriptMatches[1].trim() : '';

    let styleMatches = styleRegex.exec(storyHtml);
    let styleCode = styleMatches && styleMatches.length > 1 ? styleMatches[1].trim() : '';

    console.log(story);
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

    console.log(editor.files);

    editor.addEventListener('fileUpdated', () => {
      story.innerHTML = editor.files.html + `<style>${editor.files.css}</style>`;
      eval(editor.files.js);
    });

    const root = document.createElement('div');
    //const separator = document.createElement('div');

    //setupDrag(story, separator, editor, () => editor.layout());

    root.className = 'story-mgt-root';
    story.className = 'story-mgt-preview';
    //separator.className = 'story-mgt-separator';
    editor.className = 'story-mgt-editor';

    root.appendChild(story);
    //root.appendChild(separator);
    root.appendChild(editor);

    let css = `

html,
body,
.story-mgt-root,
#root {
  height: 100%;
  width: 100%;
}

body 
{ 
  overflow: hidden;
  margin: 0px;
}

.story-mgt-root
{
  display: flex;
  flex-direction: row;
}

.story-mgt-preview
{
  width: 50%;
  overflow: auto;
  min-width: 200px;
  min-height: 150px;
}

.story-mgt-editor {
  width: 50%;
  min-width: 200px;
  min-height: 150px;
}

.story-mgt-separator {
  cursor: col-resize;
  background-color: #aaa;
  background-image: url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='10' height='30'><path d='M2 0 v30 M5 0 v30 M8 0 v30' fill='none' stroke='black'/></svg>");
  background-repeat: no-repeat;
  background-position: center;
  width: 10px;
  height: 100%;

  /* prevent browser's built-in drag from interfering */
  -moz-user-select: none;
  -ms-user-select: none;
  user-select: none;
}

@media (max-width: 800px) {
  .story-mgt-root
  {
    display: flex;
    flex-direction: column;
  }
  
  .story-mgt-preview
  {
    height: 50%;
    width: unset;
  }
  
  .story-mgt-editor {
    height: 50%;
    width: unset;
  }

  .story-mgt-separator {
    cursor: row-resize;
  // background-image: url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='10' height='30'><path d='M2 0 v30 M5 0 v30 M8 0 v30' fill='none' stroke='black'/></svg>");
    background-repeat: no-repeat;
    background-position: center;
    width: 100%;
    height: 10px;
  }
}
    `;

    let style = document.createElement('style');
    style.type = 'text/css';
    style.appendChild(document.createTextNode(css));

    let head = document.head || document.getElementsByTagName('head')[0];
    head.appendChild(style);

    return root;
  }
});
