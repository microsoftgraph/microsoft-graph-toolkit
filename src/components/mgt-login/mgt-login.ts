import { LitElement, html, customElement, property } from "lit-element";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";

import { Providers } from "../../Providers";
import { styles } from "./mgt-login-css";

import { MgtPersonDetails } from "../mgt-person/mgt-person";
import "../mgt-person/mgt-person";

@customElement("mgt-login")
export class MgtLogin extends LitElement {
  private _loginButtonRect: ClientRect;
  private _popupRect: ClientRect;

  @property({ attribute: false }) private _showMenu: boolean = false;
  @property({ attribute: false }) private _user: MicrosoftGraph.User;

  @property({
    attribute: "user-details",
    type: Object
  })
  userDetails: MgtPersonDetails;

  static get styles() {
    return styles;
  }

  constructor() {
    super();
    Providers.onProvidersChanged(_ => this.init());
    this.init();
  }

  private fireCustomEvent(eventName: string): boolean {
    let event = new CustomEvent(eventName, {
      cancelable: true,
      bubbles: false
    });
    return this.dispatchEvent(event);
  }

  updated(changedProps) {
    if (changedProps.get("_showMenu") === false) {
      // get popup bounds
      const popup = this.shadowRoot.querySelector(".popup");
      if (popup) {
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
              transformOrigin: "top left",
              transform: `
              translate(${deltaX}px, ${deltaY}px)
              scale(${deltaW}, ${deltaH})
            `,
              backgroundColor: `#eaeaea`
            },
            {
              transformOrigin: "top left",
              transform: "none",
              backgroundColor: `white`
            }
          ],
          {
            duration: 300,
            easing: "ease-in-out",
            fill: "both"
          }
        );
      }
    } else if (changedProps.get("_showMenu") === true) {
      // get login button bounds
      const loginButton = this.shadowRoot.querySelector(".login-button");
      if (loginButton) {
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
              transformOrigin: "top left",
              transform: `
              translate(${deltaX}px, ${deltaY}px)
              scale(${deltaW}, ${deltaH})
            `
            },
            {
              transformOrigin: "top left",
              transform: "none"
            }
          ],
          {
            duration: 200,
            easing: "ease-out",
            fill: "both"
          }
        );
      }
    }
  }

  firstUpdated() {
    window.onclick = (event: any) => {
      if (event.target !== this) {
        // get popup bounds
        const popup = this.shadowRoot.querySelector(".popup");
        if (popup){
          this._popupRect = popup.getBoundingClientRect();
          this._showMenu = false;
        }
      }
    };
  }

  private async init() {
    const provider = Providers.GlobalProvider;
    if (provider) {
      provider.onLoginChanged(_ => this.loadState());
      await this.loadState();
    }
  }

  private async loadState() {
    if (this.userDetails) {
      this._user = null;
      return;
    }

    const provider = Providers.GlobalProvider;

    if (provider && provider.isLoggedIn) {
      this._user = await provider.graph.me();
    } else {
      this._user = null;
    }
  }

  private onClick() {
    if (this._user || this.userDetails) {
      // get login button bounds
      const loginButton = this.shadowRoot.querySelector(".login-button");
      if (loginButton) {
        this._loginButtonRect = loginButton.getBoundingClientRect();
        this._showMenu = !this._showMenu;
      }
    } else {
      if (this.fireCustomEvent("loginInitiated")) {
        this.login();
      }
    }
  }

  public async login() {
    if (this.userDetails) {
      return;
    }

    const provider = Providers.GlobalProvider;

    if (provider && provider.login) {
      await provider.login();

      if (provider.isLoggedIn) {
        this.fireCustomEvent("loginCompleted");
      } else {
        this.fireCustomEvent("loginFailed");
      }

      await this.loadState();
    }
  }

  public async logout() {
    if (!this.fireCustomEvent("logoutInitiated")) {
      return;
    }

    if (this.userDetails) {
      this.userDetails = null;
      return;
    }

    const provider = Providers.GlobalProvider;
    if (provider && provider.logout) {
      await provider.logout();
      this.fireCustomEvent("logoutCompleted");
    }

    this._showMenu = false;
  }

  render() {
    const content =
      this._user || this.userDetails
        ? this.renderLoggedIn()
        : this.renderLogIn();

    return html`
      <div class="root">
        <button class="login-button" @click=${this.onClick}>
          ${content}
        </button>
        ${this.renderMenu()}
      </div>
    `;
  }

  renderLogIn() {
    return html`
      <i class="login-icon ms-Icon ms-Icon--Contact"></i>
      <span>
        Sign In
      </span>
    `;
  }

  renderLoggedIn() {
    if (this._user) {
      return html`
        <mgt-person person-query="me" show-name />
      `;
    } else if (this.userDetails) {
      return html`
        <mgt-person
          person-details=${JSON.stringify(this.userDetails)}
          show-name
        />
      `;
    } else {
      return this.renderLogIn();
    }
  }

  renderMenu() {
    if (!this._user && !this.userDetails) {
      return;
    }

    let personComponent = this._user
      ? html`
          <mgt-person person-query="me" show-name show-email />
        `
      : html`
          <mgt-person
            person-details=${JSON.stringify(this.userDetails)}
            show-name
            show-email
          />
        `;

    return html`
      <div class="popup ${this._showMenu ? "show-menu" : ""}">
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
}
