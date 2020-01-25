import addons, { makeDecorator } from '@storybook/addons';
import * as monaco from 'monaco-editor/esm/vs/editor/editor.api';

function debounce(func, wait, immediate) {
  var timeout;
  return function() {
    var context = this,
      args = arguments;
    var later = function() {
      timeout = null;
      if (!immediate) func.apply(context, args);
    };
    var callNow = immediate && !timeout;
    clearTimeout(timeout);
    timeout = setTimeout(later, wait);
    if (callNow) func.apply(context, args);
  };
}

export const withCode = makeDecorator({
  name: `withCode`,
  parameterName: 'myParameter',
  skipIfNoParametersOrOptions: false,
  wrapper: (getStory, context, { parameters }) => {
    let story = getStory(context);

    let editorRoot = document.createElement('div');
    editorRoot.style.height = '200px';
    editorRoot.style.width = '100%;';
    editorRoot.style.flexBasis = '200px';
    editorRoot.style.flexShrink = '0';

    let editor = monaco.editor.create(editorRoot, {
      value: story.innerHTML.replace(/\n?<!---->\n?/g, ''),
      language: 'html',
      theme: 'vs-dark',
      scrollBeyondLastLine: false,
      minimap: {
        enabled: false
      }
    });

    editor.changeViewZones(changeAccessor => {
      const domNode = document.createElement('div');
      const viewZoneId = changeAccessor.addZone({
        afterLineNumber: 0,
        heightInLines: 1,
        domNode: domNode
      });
    });

    const updateStory = debounce(() => {
      story.innerHTML = editor.getValue();
    }, 500);

    editor.getModel().onDidChangeContent(event => {
      updateStory();
    });

    window.addEventListener('resize', () => {
      editor.layout();
    });

    setTimeout(_ => {
      editor.layout();
    }, 2);

    const root = document.createElement('div');
    root.style.height = '100vh';
    root.style.display = 'flex';
    root.style.flexDirection = 'column';

    story.style.flexGrow = '1';
    story.style.overflow = 'auto';

    root.appendChild(story);
    root.appendChild(editorRoot);

    let css = `
    body 
    { overflow: hidden; }`;
    let head = document.head || document.getElementsByTagName('head')[0];
    let style = document.createElement('style');

    head.appendChild(style);

    style.type = 'text/css';
    style.appendChild(document.createTextNode(css));

    return root;
  }
});
