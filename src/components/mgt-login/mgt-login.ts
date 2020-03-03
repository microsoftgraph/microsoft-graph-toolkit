/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { User } from '@microsoft/microsoft-graph-types';
import { customElement, html, property } from 'lit-element';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import '../../styles/fabric-icon-font';
import { MgtBaseComponent } from '../baseComponent';
import '../mgt-person/mgt-person';
import { styles } from './mgt-login-css';

/**
 * Web component button and flyout control to facilitate Microsoft identity platform authentication
 *
 * @export
 * @class MgtLogin
 * @extends {MgtBaseComponent}
 *
 * @cssprop --font-size - {Length} Login font size
 * @cssprop --font-weight - {Length} Login font weight
 * @cssprop --height - {String} Login height percentage
 * @cssprop --margin - {String} Margin size
 * @cssprop --padding - {String} Padding size
 * @cssprop --color - {Color} Login font color
 * @cssprop --background-color - {Color} Login background color
 * @cssprop --background-color--hover - {Color} Login background hover color
 * @cssprop --popup-content-background-color - {Color} Popup content background color
 * @cssprop --popup-command-font-size - {Length} Popup command font size
 */
@customElement('mgt-login')
export class MgtLogin extends MgtBaseComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * allows developer to use specific user details for login
   * @type {MgtPersonDetails}
   */
  @property({
    attribute: 'user-details',
    type: Object
  })
  public userDetails: User;

  private _image: string;

  /**
   * determines if login menu popup should be showing
   * @type {boolean}
   */
  @property({ attribute: false }) private _showMenu: boolean;

  constructor() {
    super();
    this.handleWindowClick = this.handleWindowClick.bind(this);
  }

  /**
   * Invoked each time the custom element is appended into a document-connected element
   *
   * @memberof MgtLogin
   */
  public connectedCallback() {
    super.connectedCallback();
    this.addEventListener('click', e => e.stopPropagation());
    window.addEventListener('click', this.handleWindowClick);
  }

  /**
   * Invoked each time the custom element is disconnected from the document's DOM
   *
   * @memberof MgtLogin
   */
  public disconnectedCallback() {
    window.removeEventListener('click', this.handleWindowClick);
    super.disconnectedCallback();
  }

  /**
   * Initiate login
   *
   * @returns {Promise<void>}
   * @memberof MgtLogin
   */
  public async login(): Promise<void> {
    if (this.userDetails) {
      return;
    }

    const provider = Providers.globalProvider;

    if (provider && provider.login) {
      await provider.login();

      if (provider.state === ProviderState.SignedIn) {
        this.fireCustomEvent('loginCompleted');
      } else {
        this.fireCustomEvent('loginFailed');
      }
    }
  }

  /**
   *
   * Initiate logout
   * @returns {Promise<void>}
   * @memberof MgtLogin
   */
  public async logout(): Promise<void> {
    if (!this.fireCustomEvent('logoutInitiated')) {
      return;
    }

    const provider = Providers.globalProvider;
    if (provider && provider.logout) {
      await provider.logout();
      this.userDetails = null;
      this._showMenu = false;
      this.fireCustomEvent('logoutCompleted');
    }
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    const content = this.userDetails ? this.renderLoggedIn() : this.renderLogIn();

    return html`
      <div class="root">
        <button ?disabled="${this.isLoadingState}" class="login-button" @click=${this.onClick} role="button">
          ${content}
        </button>
        ${this.renderMenu()}
      </div>
    `;
  }

  /**
   * Load state into the component.
   *
   * @protected
   * @returns
   * @memberof MgtLogin
   */
  protected async loadState() {
    const provider = Providers.globalProvider;
    if (provider) {
      if (provider.state === ProviderState.SignedIn) {
        const batch = provider.graph.forComponent(this).createBatch();
        batch.get('me', 'me', ['user.read']);
        batch.get('photo', 'me/photo/$value', ['user.read']);
        const response = await batch.execute();

        this._image = response.photo;
        this.userDetails = response.me;
      } else if (provider.state === ProviderState.SignedOut) {
        this.userDetails = null;
      } else {
        // Loading
        this._showMenu = false;
        return;
      }
    }
  }

  private handleWindowClick(e: MouseEvent) {
    this._showMenu = false;
  }

  private renderLogIn() {
    return html`
      <i class="login-icon ms-Icon ms-Icon--Contact"></i>
      <span aria-label="Sign In">
        Sign In
      </span>
    `;
  }

  private renderLoggedIn() {
    if (this.userDetails) {
      return html`
        <mgt-person .personDetails=${this.userDetails} .personImage=${this._image} show-name />
      `;
    } else {
      return this.renderLogIn();
    }
  }

  private renderMenu() {
    if (!this.userDetails) {
      return;
    }

    const personComponent = html`
      <mgt-person .personDetails=${this.userDetails} .personImage=${this._image} show-name show-email />
    `;

    return html`
      <mgt-flyout .isOpen=${this._showMenu}>
        <div slot="flyout" class="flyout">
          <div class="popup">
            <div class="popup-content">
              <div>
                ${personComponent}
              </div>
              <div class="popup-commands">
                <ul>
                  <li>
                    <button class="popup-command" @click=${this.logout} aria-label="Sign Out">
                      Sign Out
                    </button>
                  </li>
                </ul>
              </div>
            </div>
          </div>
        </div>
      </mgt-flyout>
    `;
  }

  private onClick(event: MouseEvent) {
    if (this.userDetails) {
      this._showMenu = !this._showMenu;
    } else {
      if (this.fireCustomEvent('loginInitiated')) {
        this.login();
      }
    }
  }
}
