import { LitElement, html, customElement, property } from 'lit-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import { Providers } from '../../providers/Providers';
import styles from './mgt-login.scss';

import '../mgt-person/mgt-person';

@customElement('mgt-login')
export class MgtLogin extends LitElement {
  @property({ attribute: false }) private _user: MicrosoftGraph.User;
  @property({ attribute: false })
  private _showMenu: boolean = false;
  private _loginButtonRect: ClientRect;
  private _popupRect: ClientRect;

  static get styles() {
    return styles;
  }

  constructor() {
    super();
    Providers.onProvidersChanged(_ => this.init());
    this.init();
  }

  updated(changedProps) {
    if (changedProps.get('_showMenu') === false) {
      // get popup bounds
      const popup = this.shadowRoot.querySelector('.popup');
      this._popupRect = popup.getBoundingClientRect();
      // console.log('last', this._popupRect);

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
            backgroundColor: `#eaeaea`
          },
          {
            transformOrigin: 'top left',
            transform: 'none',
            backgroundColor: `white`
          }
        ],
        {
          duration: 300,
          easing: 'ease-in-out',
          fill: 'both'
        }
      );
    } else if (changedProps.get('_showMenu') === true) {
      // get login button bounds
      const loginButton = this.shadowRoot.querySelector('.login-button');
      this._loginButtonRect = loginButton.getBoundingClientRect();
      // console.log('last', this._loginButtonRect);

      // invert variables
      const deltaX = this._popupRect.left - this._loginButtonRect.left;
      const deltaY = this._popupRect.top - this._loginButtonRect.top;
      const deltaW = this._popupRect.width / this._loginButtonRect.width;
      const deltaH = this._popupRect.height / this._loginButtonRect.height;

      // play back
      loginButton.animate(
        [
          {
            transformOrigin: 'top left',
            transform: `
               translate(${deltaX}px, ${deltaY}px)
               scale(${deltaW}, ${deltaH})
             `
          },
          {
            transformOrigin: 'top left',
            transform: 'none'
          }
        ],
        {
          duration: 200,
          easing: 'ease-out',
          fill: 'both'
        }
      );
    }
  }

  firstUpdated() {
    window.onclick = (event: any) => {
      if (event.target !== this) {
        // get popup bounds
        const popup = this.shadowRoot.querySelector('.popup');
        this._popupRect = popup.getBoundingClientRect();
        // console.log('first', this._popupRect);

        this._showMenu = false;
      }
    };
  }

  private async init() {
    const provider = Providers.getAvailable();
    if (provider) {
      provider.onLoginChanged(_ => this.loadState());
      await this.loadState();
    }
  }

  private async loadState() {
    const provider = Providers.getAvailable();

    if (provider && provider.isLoggedIn) {
      this._user = await provider.graph.me();
    }
  }

  private clicked() {
    if (this._user) {
      // get login button bounds
      const loginButton = this.shadowRoot.querySelector('.login-button');
      this._loginButtonRect = loginButton.getBoundingClientRect();
      // console.log('first', this._loginButtonRect);

      this._showMenu = true;
    } else {
      this.login();
    }
  }

  public async login() {
    const provider = Providers.getAvailable();

    if (provider) {
      await provider.login();
      await this.loadState();
    }
  }

  public async logout() {
    const provider = Providers.getAvailable();
    if (provider) {
      await provider.logout();
    }
  }

  render() {
    const content = this._user ? this.renderLoggedIn() : this.renderLoggedOut();

    return html`
      <div class="root">
        <button class="login-button" @click=${this.clicked}>
          ${content}
        </button>
        ${this.renderMenu()}
      </div>
    `;
  }

  renderLoggedOut() {
    return html`
      <i class="login-icon ms-Icon ms-Icon--Contact"></i>
      <span>
        Sign In
      </span>
    `;
  }

  renderLoggedIn() {
    return html`
      <mgt-person person-query="me" show-name />
    `;
  }

  renderMenu() {
    if (!this._user) {
      return;
    }

    return html`
      <div class="popup ${this._showMenu ? 'show-menu' : ''}">
        <div class="popup-content">
          <div>
            <mgt-person person-query="me" show-name show-email />
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
}
