import { LitElement, customElement, html, property } from 'lit-element';
import { styles } from './mgt-arrow-options-css';

@customElement('mgt-arrow-options')
export class MgtArrowOptions extends LitElement {
  public static get styles() {
    return styles;
  }

  @property({ type: Boolean })
  public open: boolean = false;

  @property({ type: String })
  public value: string = '';

  @property({ type: Object })
  public options: { [name: string]: (e: MouseEvent) => any | void } = {};

  @property({ attribute: 'read-only', type: Boolean })
  public readOnly: boolean = false;

  public constructor() {
    super();

    window.addEventListener('click', () => {
      this.open = false;
    });
  }

  public onHeaderClick(e: MouseEvent) {
    e.preventDefault();
    e.stopPropagation();
    this.open = !this.open;
  }

  public render() {
    let opts = this.getMenuOptions();

    let arrowIcon =
      Object.keys(this.options).length > 1
        ? html`
            <span class="ArrowIcon" @click=${e => this.onHeaderClick(e)}>
              \uE70D
            </span>
          `
        : null;

    return html`
      <span class="Header">
        <span class="CurrentValue">${this.value}</span>
        ${arrowIcon}
      </span>
      <div class="Menu ${this.open ? 'Open' : 'Closed'}">
        ${opts}
      </div>
    `;
  }

  private getMenuOptions() {
    let keys = Object.keys(this.options);
    let funcs = this.options;

    return keys.map(
      opt => html`
        <div
          class="MenuOption"
          @click="${(e: MouseEvent) => {
            this.open = false;
            funcs[opt](e);
          }}"
        >
          <span class="MenuOptionCheck ${this.value === opt ? 'CurrentValue' : ''}">
            \uE73E
          </span>
          <span class="MenuOptionName">${opt}</span>
        </div>
      `
    );
  }
}
