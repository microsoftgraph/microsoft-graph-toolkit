import addons, { makeDecorator } from '@storybook/addons';
import './components/tab';
import { TabElement } from './components/tab';

let debounce = (func, wait, immediate) => {
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
    let scriptCode = scriptMatches.length > 1 ? scriptMatches[1].trim() : null;

    let styleMatches = styleRegex.exec(storyHtml);
    let styleCode = styleMatches.length > 1 ? styleMatches[1].trim() : null;

    storyHtml = storyHtml
      .replace(styleRegex, '')
      .replace(scriptRegex, '')
      .replace(/\n?<!---->\n?/g, '')
      .trim();

    let editor = new TabElement();

    editor.files = {
      html: storyHtml,
      js: scriptCode,
      css: styleCode
    };

    editor.addEventListener('fileUpdated', () => {
      story.innerHTML = editor.files.html + `<style>${editor.files.css}</style>`;
      eval(editor.files.js);
    });

    const root = document.createElement('div');
    root.style.height = '100vh';
    root.style.display = 'flex';
    root.style.flexDirection = 'column';

    story.style.flexGrow = '1';
    story.style.overflow = 'auto';

    root.appendChild(story);
    root.appendChild(editor);

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
