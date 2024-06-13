import { LitElement, css, html } from 'lit';
import * as monaco from 'monaco-editor/esm/vs/editor/editor.api';
import { generateProject } from './projectBuilder';

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
        font-weight: 400;
        padding: 8px 18px;
        display: inline-block;
        cursor: pointer;
        user-select: none;
        margin: 0px -2px 0px 0px;
        border: 1px solid transparent;
      }

      .tab-right {
        float: right;
        @media (forced-colors: active) and (prefers-color-scheme: dark) {
          svg,
          svg>path {
            fill: rgb(255, 255, 255);
            color: rgb(255, 255, 255);
          }
        }

        @media (forced-colors: active) and (prefers-color-scheme: light) {
          svg,
          svg>path {
            fill: rgb(0, 0, 0);
            color: rgb(0, 0, 0);
          }
        }
      }

      .tab[aria-selected='true'] {
        background-color: white;
        color: rgb(51, 51, 51);
        font-weight: 600;
        border: 2px solid transparent;
        text-decoration: underline;
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

  constructor(fileTypes) {
    super();
    this.internalFiles = [];
    this.fileTypes = fileTypes ?? ['html', 'js', 'css'];

    this.editorRoot = document.createElement('div');
    this.editorRoot.setAttribute('slot', 'editor');
    this.editorRoot.style.height = '100%';
    this.tabFocus = 0;

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
      },
      react: {
        model: monaco.editor.createModel('', 'typescript'),
        state: null
      }
    };

    this.currentEditorState = this.editorState[this.fileTypes[0]];

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
    this.editor.dispose();
    window.removeEventListener('resize', this.handleResize);
  }

  showTab(type) {
    this.editor.updateOptions({ readOnly: type === 'react' });

    this.currentType = type;
    if (this.files && typeof this.files[type] !== 'undefined') {
      this.currentEditorState.state = this.editor.saveViewState();

      this.currentEditorState = this.editorState[type];
      this.editor.setModel(this.currentEditorState.model);
      this.editor.restoreViewState(this.currentEditorState.state);
    }
  }

  generateStackblitzProject = () => {
    this.currentType = 'stackblitz';
    generateProject(this.title, this.files);
  };

  tabKeyDown = e => {
    const tabs = this.renderRoot.querySelectorAll('.tab');
    // Move right
    if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') {
      tabs[this.tabFocus].setAttribute('tabindex', -1);
      if (e.key === 'ArrowRight') {
        this.tabFocus++;
        // If we're at the end, go to the start
        if (this.tabFocus >= tabs.length) {
          this.tabFocus = 0;
        }
        // Move left
      } else if (e.key === 'ArrowLeft') {
        this.tabFocus--;
        // If we're at the start, move to the end
        if (this.tabFocus < 0) {
          this.tabFocus = tabs.length - 1;
        }
      }

      tabs[this.tabFocus].setAttribute('tabindex', 0);
      tabs[this.tabFocus].focus();
    }
  };

  render() {
    return html`
      <div class="root">
        <div class="tab-root" role="tablist" tabindex="0" @keydown="${this.tabKeyDown}" >
          ${this.fileTypes.map(
            type => html`
              <button
                tabindex="${type === this.currentType ? 0 : -1}"
                @click="${_ => this.showTab(type)}"
                id="${type}"
                role="tab"
                class="tab"
                aria-selected="${type === this.currentType}"
                aria-controls="${`tab-${type}`}"
              >
                ${type}
              </button>
            `
          )}

          <button
            @click="${this.generateStackblitzProject}"
            tabindex="${this.currentType === 'stackblitz' ? 0 : -1}"
            id="project"
            role="tab"
            class="tab tab-right"
            aria-selected="${this.currentType === 'stackblitz'}"
            title="Edit in StackBlitz"
          >
            <svg viewBox="0 0 14 14" width="14px" height="14px" class="css-149xqrd"><path d="M2 1a1 1 0 0 0-1 1v10a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1V7.5a.5.5 0 0 0-1 0V12H2V2h4.5a.5.5 0 0 0 0-1H2Z"></path><path d="M7.35 7.36 12 2.7v1.8a.5.5 0 0 0 1 0v-3a.5.5 0 0 0-.5-.5h-3a.5.5 0 1 0 0 1h1.8L6.64 6.64a.5.5 0 1 0 .7.7Z"></path></svg>
          </button>

        </div>
        <div
          class="editor-root"
          role="tabpanel"
          id="${`tab-${this.currentType}`}"
          aria-labelledby="${this.currentType}"
          tabindex=0
        >
          <slot name="editor"></slot>
        </div>
        ${this.fileTypes.map(type =>
          type !== this.currentType
            ? html`
            <div
              role="tabpanel"
              id="${`tab-${type}`}"
              aria-labelledby="${type}"
              tabindex=0
              hidden
            ></div>
          `
            : ''
        )}
      </div>
    `;
  }
}

customElements.define('mgt-sb-editor', EditorElement);
