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

@customElement('mgt-login')
export class MgtLogin extends MgtBaseComponent {
  static get styles() {
    return styles;
  }

  @property({
    attribute: 'user-details',
    type: Object
  })
  public userDetails: User;

  private _image: string;
  private _loginButtonRect: ClientRect;
  private _popupRect: ClientRect;
  private _openLeft: boolean = false;

  @property({ attribute: false }) private _showMenu: boolean = false;
  @property({ attribute: false }) private _loading: boolean = true;

  constructor() {
    super();
    Providers.onProviderUpdated(() => this.loadState());
    this.loadState();
  }

  public updated(changedProps) {
    if (changedProps.get('_showMenu') === false) {
      // get popup bounds
      const popup = this.renderRoot.querySelector('.popup');
      if (popup && popup.animate) {
        this._popupRect = popup.getBoundingClientRect();

        // invert variables
        const deltaX = this._loginButtonRect.left - this._popupRect.left;
        const deltaY = this._loginButtonRect.top - this._popupRect.top;
        const deltaW = this._loginButtonRect.width / this._popupRect.width;
        const deltaH = this._loginButtonRect.height / this._popupRect.height;

        // play back
        popup.animate(
          [
            {
              transformOrigin: 'top left',
              transform: `
              translate(${deltaX}px, ${deltaY}px)
              scale(${deltaW}, ${deltaH})
            `,
              backgroundColor: '#eaeaea'
            },
            {
              transformOrigin: 'top left',
              transform: 'none',
              backgroundColor: 'white'
            }
          ],
          {
            duration: 100,
            easing: 'ease-in-out',
            fill: 'both'
          }
        );
      }
    }
  }

  public firstUpdated() {
    window.addEventListener('click', (event: MouseEvent) => {
      // get popup bounds
      const popup = this.renderRoot.querySelector('.popup');
      if (popup) {
        this._popupRect = popup.getBoundingClientRect();
        this._showMenu = false;
      }
    });
  }

  public async login() {
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

      await this.loadState();
    }
  }

  public async logout() {
    if (!this.fireCustomEvent('logoutInitiated')) {
      return;
    }

    if (this.userDetails) {
      this.userDetails = null;
      return;
    }

    const provider = Providers.globalProvider;
    if (provider && provider.logout) {
      await provider.logout();
      this.fireCustomEvent('logoutCompleted');
    }

    this._showMenu = false;
  }

  public render() {
    const content = this.userDetails ? this.renderLoggedIn() : this.renderLogIn();

    return html`
      <div class="root">
        <button ?disabled="${this._loading}" class="login-button" @click=${this.onClick}>
          ${content}
        </button>
        ${this.renderMenu()}
      </div>
    `;
  }

  public renderLogIn() {
    return html`
      <i class="login-icon ms-Icon ms-Icon--Contact"></i>
      <span>
        Sign In
      </span>
    `;
  }

  public renderLoggedIn() {
    if (this.userDetails) {
      return html`
        <mgt-person .personDetails=${this.userDetails} .personImage=${this._image} show-name />
      `;
    } else {
      return this.renderLogIn();
    }
  }

  public renderMenu() {
    if (!this.userDetails) {
      return;
    }

    const personComponent = html`
      <mgt-person .personDetails=${this.userDetails} .personImage=${this._image} show-name show-email />
    `;

    return html`
      <div class="popup ${this._openLeft ? 'open-left' : ''} ${this._showMenu ? 'show-menu' : ''}">
        <div class="popup-content">
          <div>
            ${personComponent}
          </div>
          <div class="popup-commands">
            <ul>
              <li>
                <button class="popup-command" @click=${this.logout}>
                  Sign Out
                </button>
              </li>
            </ul>
          </div>
        </div>
      </div>
    `;
  }

  private async loadState() {
    if (this.userDetails) {
      return;
    }

    const provider = Providers.globalProvider;

    if (provider) {
      this._loading = true;
      if (provider.state === ProviderState.SignedIn) {
        const batch = provider.graph.createBatch();
        batch.get('me', 'me', ['user.read']);
        batch.get('photo', 'me/photo/$value', ['user.read']);
        const response = await batch.execute();

        this._image = response.photo;
        this.userDetails = response.me;
      } else if (provider.state === ProviderState.SignedOut) {
        this.userDetails = null;
      } else {
        return;
      }
    }

    this._loading = false;
  }

  private onClick(event: MouseEvent) {
    event.stopPropagation();
    if (this.userDetails) {
      // get login button bounds
      const loginButton = this.renderRoot.querySelector('.login-button');
      if (loginButton) {
        this._loginButtonRect = loginButton.getBoundingClientRect();

        const leftEdge = this._loginButtonRect.left;
        const rightEdge = (window.innerWidth || document.documentElement.clientWidth) - this._loginButtonRect.right;
        this._openLeft = rightEdge < leftEdge;

        this._showMenu = !this._showMenu;
      }
    } else {
      if (this.fireCustomEvent('loginInitiated')) {
        this.login();
      }
    }
  }
}
