import { LitElement, html, customElement, property } from 'lit-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import { Providers } from '../../providers';
import { style } from './graph-login.style';

import '../graph-persona/graph-persona';

@customElement('graph-login')
export class LoginComponent extends LitElement {
  @property({ attribute: false }) private _user: MicrosoftGraph.User;
  @property({ attribute: false }) private _showMenu: boolean = false;

  static get styles() {
    return style;
  }

  constructor() {
    super();
    Providers.onProvidersChanged(_ => this.init());
    this.init();
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

  private clicked() {
    if (this._user) {
      this._showMenu = !this._showMenu;
    } else {
      this.login();
    }
  }

  render() {
    let content = this._user ? this.renderLoggedIn() : this.renderLoggedOut();

    return html`
      <div class="login-root">
        <button class="login-root-button" @click=${this.clicked}>
          ${content}
        </button>
        ${this.renderMenu()}
      </div>
    `;
  }

  renderLoggedOut() {
    return html`
      <div class="login-signed-out-root">
        <div>
          <i class="ms-Icon ms-Icon--AddFriend"></i>
        </div>
        <div class="login-signed-out-content">
          Sign In
        </div>
      </div>
    `;
  }

  renderLoggedIn() {
    return html`
      <div>
        <div class="login-signed-in-root">
          <div class="login-signed-in-image">
            <graph-persona persona-id="me" image-size="24" />
          </div>
          <div class="login-signed-in-content">
            ${this._user.displayName}
          </div>
        </div>
      </div>
    `;
  }

  renderMenu() {
    if (!this._user) {
      return;
    }

    return html`
      <div class="login-menu-root ${this._showMenu ? 'show-menu' : ''}">
        <div class="login-menu-beak"></div>
        <div class="login-menu-beak-cover"></div>
        <div class="login-menu-content">
          <div class="login-menu-user-profile">
            <div class="login-menu-user-image">
              <graph-persona persona-id="me" image-size="65" />
            </div>
            <div class="login-menu-user-details">
              <div class="login-menu-user-display-name">
                ${this._user.displayName}
              </div>
              <div class="login-menu-user-email">${this._user.mail}</div>
            </div>
          </div>
          <div class="login-menu-commands">
            <ul>
              <li>
                <button class="login-menu-command" @click=${this.logout}>
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
