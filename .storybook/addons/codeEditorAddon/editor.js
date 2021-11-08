import { LitElement, css, html, property, customElement } from 'lit-element';
import * as monaco from 'monaco-editor/esm/vs/editor/editor.api';

let debounce = (func, wait, immediate) => {
  var timeout;
  return function () {
    var context = this,
      args = arguments;
    var later = function () {
      timeout = null;
      if (!immediate) func.apply(context, args);
    };
    var callNow = immediate && !timeout;
    clearTimeout(timeout);
    timeout = setTimeout(later, wait);
    if (callNow) func.apply(context, args);
  };
};

export class EditorElement extends LitElement {
  static get styles() {
    return css`
      :host {
        display: block;
        height: 100%;
      }

      .root {
        height: 100%;
        display: flex;
        flex-direction: column;
        background-color: rgb(243, 243, 243);
      }

      .editor-root {
        flex-basis: 100%;
      }

      .tab {
        background-color: rgb(236, 236, 236);
        color: #616161;
        font-family: -apple-system, BlinkMacSystemFont, sans-serif;
        font-size: 11px;
        padding: 8px 18px;
        display: inline-block;
        cursor: pointer;
        user-select: none;
        margin: 0px -2px 0px 0px;
      }

      .tab.selected {
        background-color: white;
        color: rgb(51, 51, 51);
        font-weight: 400;
      }
    `;
  }

  static get properties() {
    return {
      files: { type: Array, attribute: false },
      currentType: { type: String }
    };
  }

  set files(value) {
    const oldValue = this.files;
    this.internalFiles = value;

    for (let type of this.fileTypes) {
      this.editorState[type].model.setValue(this.files[type] + '\n');
    }

    this.showTab(this.currentType);
    this.requestUpdate('files', oldValue);
  }

  get files() {
    return this.internalFiles;
  }

  constructor() {
    super();
    this.internalFiles = [];
    this.fileTypes = ['html', 'js', 'css'];

    this.editorRoot = document.createElement('div');
    this.editorRoot.setAttribute('slot', 'editor');
    this.editorRoot.style.height = '100%';

    this.updateCurrentFile = debounce(() => {
      this.files[this.currentType] = this.editor.getValue();
      let event = new CustomEvent('fileUpdated');
      this.dispatchEvent(event);
    }, 1000);

    this.setupEditor(this.editorRoot);
    this.appendChild(this.editorRoot);

    this.handleResize = this.handleResize.bind(this);
    this.showTab(this.fileTypes[0]);
  }

  setupEditor(htmlElement) {
    this.editorState = {
      js: {
        model: monaco.editor.createModel('', 'javascript'),
        state: null
      },
      css: {
        model: monaco.editor.createModel('', 'css'),
        state: null
      },
      html: {
        model: monaco.editor.createModel('', 'html'),
        state: null
      }
    };

    this.currentEditorState = this.editorState.html;

    this.editor = monaco.editor.create(htmlElement, {
      model: this.currentEditorState.model,
      scrollBeyondLastLine: false,
      readOnly: true,
      minimap: {
        enabled: false
      }
    });

    // Exit the current editor
    this.editor.addCommand(monaco.KeyCode.Escape, () => {
      this.editor.updateOptions({ readOnly: true });
      this.shadowRoot.getElementById(this.currentType).focus();
    });

    const changeViewZones = () => {
      this.editor.changeViewZones(changeAccessor => {
        const domNode = document.createElement('div');
        changeAccessor.addZone({
          afterLineNumber: 0,
          heightInLines: 1,
          domNode: domNode
        });
      });
    };
    this.editor.onDidChangeModel(changeViewZones);
    changeViewZones();

    this.editor.onDidChangeModelContent(() => {
      this.updateCurrentFile();
    });
  }

  layout() {
    this.editorRoot.style.height = `calc(${this.style.height} - 38px)`;
    this.editor.layout();
  }

  handleResize() {
    this.editorRoot.style.height = `${this.clientHeight - 38}px`;
    this.editor.layout();
  }

  connectedCallback() {
    super.connectedCallback();
    this.editor.layout();

    window.addEventListener('resize', this.handleResize);

    setTimeout(_ => {
      this.editor.layout();
    }, 2);
  }

  disconnectedCallback() {
    window.removeEventListener('resize', this.handleResize);
  }

  showTab(type, event) {
    this.editor.updateOptions({ readOnly: false });
    if (event && event.keyCode != 13) return;

    this.currentType = type;
    if (this.files && typeof this.files[type] !== 'undefined') {
      this.currentEditorState.state = this.editor.saveViewState();

      this.currentEditorState = this.editorState[type];
      this.editor.setModel(this.currentEditorState.model);
      this.editor.restoreViewState(this.currentEditorState.state);
    }

    this.editor.getAction('editor.action.formatDocument').run();
  }

  render() {
    return html`
      <div class="root" tabindex=0>
        <div class="tab-root" tabindex=0">
          ${this.fileTypes.map(
            type => html`
              <div
                @keydown=${e => this.showTab(type, e)}
                tabindex=0
                @click="${_ => this.showTab(type)}"
                id="${type}"
                class="tab ${type === this.currentType ? 'selected' : ''}">
                ${type}
              </div>
            `
          )}
        </div>
        <div class="editor-root">
          <slot name="editor"></slot>
        </div>
      </div>
    `;
  }
}

customElements.define('mgt-sb-editor', EditorElement);
