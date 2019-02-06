import { Component, State, Element, Method } from '@stencil/core';

import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"
import * as Auth from '../../auth/Auth'
import { IAuthProvider } from '../../auth/IAuthProvider';

@Component({
  tag: 'my-login',
  styleUrl: 'my-login.css',
  shadow: true
})
export class MyLogin {

  private provider : IAuthProvider;

  async componentWillLoad()
  {
    Auth.onAuthProviderChanged(_ => this.init())
    this.init();
  }

  componentDidLoad() {
    this.$dropdownElement = this.$rootElement.shadowRoot.querySelector('.login-dropdown-root');
    this.dropdownVisible = false;
  }

  componentDidUpdate() {
    this.componentDidLoad();
  }

  private async init() {
    this.provider = Auth.getAuthProvider();
    if (this.provider) {
      this.provider.onLoginChanged(_ => this.loadState());
      await this.loadState();
    }
  }

  private async loadState() {
    if (this.provider && this.provider.isLoggedIn) {
      this.user = await this.provider.graph.me();
      let profileImage = await this.provider.graph.myPhoto();
      let reader = new FileReader();
      reader.onload = () => {
        this.profileImage = reader.result;
      }
      reader.readAsDataURL(await profileImage.blob())
    }
  }

  @Method()
  async login()
  {
    if (this.provider)
    {
      await this.provider.login();
      await this.loadState();
    }
  }

  @Method()
  async logout()
  {
    if (this.provider)
    {
      await this.provider.logout();
    }
  }

  toggleDropdown() {
    if (this.$dropdownElement) {
      this.dropdownVisible = !this.dropdownVisible;
      this.$dropdownElement.className = 'login-dropdown-root' +
        (this.dropdownVisible ? ' show-dropdown' : '');
    }
  }

  clicked() {
    if (this.user) {
      this.toggleDropdown();
    } else {
      this.login();
    }
  }

  @Element() $rootElement : HTMLElement;

  private $dropdownElement : HTMLElement;
  private dropdownVisible : boolean = false;

  @State() user : MicrosoftGraph.User;
  @State() profileImage: string | ArrayBuffer;

  render() {
    let content = this.user ?
        this.renderLoggedIn() : this.renderLoggedOut();

    return <div class='login-root'>
      <button class='login-root-button' onClick={() => {this.clicked()}}>
        {content}
      </button>
      {this.renderDropDown()}
    </div>
  }

  renderLoggedOut() {
    return 'Login';
  }

  renderLoggedIn() {
    return <div>
        <div class="login-header">
          <div class="login-header-user">
            {this.user.displayName}
          </div>
          <div class="login-header-user-image-container">
            {this.renderImage()}
          </div>
        </div>
      </div>;
  }

  renderImage() {
    if (this.profileImage) {
      return <img class="login-user-image" src={this.profileImage as string}></img>
    } else {
      return <div class="login-user-no-image">TD</div>
    }
  }

  renderDropDown() {
    if (!this.user) {
      return;
    }

    return <div class="login-dropdown-root">
      <div>{this.user.displayName}</div>
      <div>{this.user.mail}</div>
      <button onClick={() => {this.logout()}}>Logout</button>
    </div>
  }
}
