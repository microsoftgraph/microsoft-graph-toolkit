import { MgtTemplatedComponent, customElement, mgtHtml } from '@microsoft/mgt-element';
import { html, nothing, TemplateResult } from 'lit';
import { property, state } from 'lit/decorators.js';
import { repeat } from 'lit/directives/repeat.js';
import { fluentAnchoredRegion, fluentButton, fluentMenu, fluentMenuItem } from '@fluentui/web-components';
import { registerFluentComponents } from '../../utils/FluentComponents';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';
import '../sub-components/mgt-flyout/mgt-flyout';
import { styles } from './mgt-menu-css';
import { MgtFlyout } from '../sub-components/mgt-flyout/mgt-flyout';

registerFluentComponents(fluentButton, fluentMenu, fluentMenuItem, fluentAnchoredRegion);

/**
 * Data type for menu items commands
 */
export type Command<T> = {
  // tslint:disable: completed-docs
  id: string;
  name: string;
  glyph?: TemplateResult;
  onClickFunction?: (e: UIEvent, item: T) => void;
  subcommands?: Command<T>[];
  // tslint:enable: completed-docs
};

/**
 * Component to provide configurable menus
 *
 * @class MgtMenu
 * @extends {MgtTemplatedComponent}
 */
@customElement('menu')
class MgtMenu<T extends { id: string }> extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * Gets the flyout element
   *
   * @protected
   * @type {MgtFlyout}
   * @memberof MgtLogin
   */
  protected get flyout(): MgtFlyout {
    return this.renderRoot.querySelector('.flyout');
  }

  /**
   * The item to which the menu is attached
   *
   * @type {T}
   * @memberof MgtMenu
   */
  @property({ attribute: false })
  public item: T;

  @state()
  private _menuOpen: boolean = false;

  /**
   * Array of items to display in the menu
   *
   * @type {{ id: string; name: string, glyph?: TemplateResult }[]}
   * @memberof MgtMenu
   */
  @property({ attribute: false })
  public commands: Command<T>[] = [];
  /**
   * Render the component
   *
   * @return {*}  {TemplateResult}
   * @memberof MgtMenu
   */
  render(): TemplateResult {
    return html`
      <fluent-button part="menu-button" id=${this.item.id} appearance="stealth" @click=${this.toggleMenu}>
        ${getSvg(SvgIcon.MoreVertical)}
      </fluent-button>
      ${this.renderMenu()}
    `;
  }

  private renderMenu() {
    return mgtHtml`
        <mgt-flyout
          class="flyout"
          part="flyout"
          light-dismiss
          @opened=${this.flyoutOpened}
          @closed=${this.flyoutClosed}
        >
          <div slot="flyout">
            <fluent-menu>
              ${this.renderMenuItems(this.commands)}
            </fluent-menu>
          </div>
        </mgt-flyout>
      `;
  }

  private flyoutOpened = () => {
    this._menuOpen = true;
  };
  private flyoutClosed = () => {
    this._menuOpen = false;
  };

  private renderMenuItems(commands: Command<T>[]): unknown {
    return repeat(
      commands,
      command => command.id,
      command => this.renderMenuItem(command)
    );
  }

  private renderMenuItem(command: Command<T>): unknown {
    return html`
          <fluent-menu-item
            @click=${(e: UIEvent) => command.onClickFunction?.(e, this.item)}
          >
            ${command.name}
            ${
              command.subcommands && command.subcommands.length > 0
                ? html`
                  <fluent-menu slot="submenu">
                    ${this.renderMenuItems(command.subcommands)}
                  </fluent-menu>
                `
                : nothing
            }
          </fluent-menu-item>`;
  }

  /**
   * Show the flyout and its content.
   *
   * @protected
   * @memberof MgtLogin
   */
  protected showFlyout(): void {
    const flyout = this.flyout;
    if (flyout) {
      flyout.open();
    }
  }

  /**
   * Dismiss the flyout.
   *
   * @protected
   * @memberof MgtLogin
   */
  protected hideFlyout(): void {
    const flyout = this.flyout;
    if (flyout) {
      flyout.close();
    }
  }

  private toggleMenu(e: Event) {
    e.stopPropagation();
    if (this._menuOpen) {
      this.hideFlyout();
    } else {
      this.showFlyout();
    }
  }
}
