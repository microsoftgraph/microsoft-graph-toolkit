import { Component, State, Element, Method } from '@stencil/core';

import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"
import { Providers, IAuthProvider} from '@m365toolkit/providers'
import { toggleClass } from '../../global/helpers';

@Component({
  tag: 'graph-login',
  styleUrl: 'graph-login.css',
  shadow: true
})
export class LoginComponent {

  @Element() $rootElement : HTMLElement;
  private $menuElement : HTMLElement;

  @State() user : MicrosoftGraph.User;
  @State() profileImage: string | ArrayBuffer;

  async componentWillLoad()
  {
    Providers.onProvidersChanged(_ => this.init())
    this.init();
  }

  componentDidLoad() {
    this.$menuElement = this.$rootElement.shadowRoot.querySelector('.login-menu-root');
  }

  componentDidUpdate() {
    this.componentDidLoad();
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
      this.user = await provider.graph.me();
      this.profileImage = await provider.graph.myPhoto();
    }
  }

  @Method()
  async login()
  {
    const provider = Providers.getAvailable();

    if (provider)
    {
      await provider.login();
      await this.loadState();
    }
  }

  @Method()
  async logout()
  {
    const provider = Providers.getAvailable();
    if (provider)
    {
      await provider.logout();
    }
  }

  toggleMenu() {
    toggleClass(this.$menuElement, 'show-menu');
  }

  clicked() {
    if (this.user) {
      this.toggleMenu();
    } else {
      this.login();
    }
  }

  render() {
    let content = this.user ?
        this.renderLoggedIn() : this.renderLoggedOut();

    return <div class='login-root'>
      <button class='login-root-button' onClick={() => {this.clicked()}}>
        {content}
      </button>
      {this.renderMenu()}
    </div>
  }

  renderLoggedOut() {
    return <div class='login-signed-out-root'>
      <div class='login-signed-out-content'>
        Sign In
      </div>
    </div>;
  }

  renderLoggedIn() {
    return <div>
        <div class="login-signed-in-root">
          <div class="login-signed-in-image">
            <graph-persona id="me" image-size="24" />
          </div>
          <div class="login-signed-in-content">
            {this.user.displayName}
          </div>
        </div>
      </div>;
  }

  renderMenu() {
    if (!this.user) {
      return;
    }

    return <div class="login-menu-root">
      <div class="login-menu-beak"></div>
      <div class="login-menu-beak-cover"></div>
      <div class="login-menu-content">
        <div class="login-menu-user-profile">
          <div class="login-menu-user-image">
            <graph-persona id="me" image-size="65" />
          </div>
          <div class="login-menu-user-details">
            <div class="login-menu-user-display-name">{this.user.displayName}</div>
            <div class="login-menu-user-email">{this.user.mail}</div>
          </div>
        </div>
        <div class='login-menu-commands'>
          <ul>
            <li>
              <button class='login-menu-command' onClick={() => {this.logout()}}>Sign Out</button>
            </li>
          </ul>
        </div>
      </div>
    </div>
  }
}
