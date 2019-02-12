import { Component, State, Element, Method } from '@stencil/core';

import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"
import { Providers, IAuthProvider} from '@m365toolkit/providers'
import { toggleClass } from '../../global/helpers';

@Component({
  tag: 'my-login',
  styleUrl: 'my-login.css',
  shadow: true
})
export class MyLogin {

  private provider : IAuthProvider;

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
    this.provider = Providers.getAvailable();
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
            {this.renderImage()}
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
            {this.renderImage()}
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

  renderImage() {
    if (this.profileImage) {
      return <img class="login-user-image" src={this.profileImage as string}></img>
    } else {
      return this.renderEmptyImage();
    }
  }

  renderEmptyImage() {
    return <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
    <mask id="mask0" mask-type="alpha" maskUnits="userSpaceOnUse" x="0" y="0" width="24" height="24">
    <circle cx="12" cy="12" r="12" fill="#FAFAFA"/>
    </mask>
    <g mask="url(#mask0)">
    <circle cx="12" cy="12" r="12" fill="#FAFAFA"/>
    <path d="M14.8326 14.4721C15.6172 14.7433 16.3239 15.124 16.9528 15.6144C17.5874 16.099 18.1239 16.6615 18.5624 17.3019C19.0066 17.9423 19.347 18.6433 19.5835 19.4048C19.8201 20.1663 19.9384 20.9596 19.9384 21.7846H18.8307C18.8307 20.8384 18.6605 19.9615 18.3201 19.1538C17.9855 18.3404 17.521 17.6365 16.9268 17.0423C16.3326 16.4481 15.6287 15.9836 14.8153 15.649C14.0076 15.3086 13.1307 15.1384 12.1845 15.1384C11.5672 15.1384 10.973 15.2163 10.4018 15.3721C9.83066 15.5279 9.29701 15.75 8.80086 16.0384C8.31047 16.3211 7.86336 16.6644 7.45951 17.0683C7.06143 17.4663 6.71816 17.9134 6.4297 18.4096C6.14701 18.9 5.92778 19.4308 5.77201 20.0019C5.61624 20.5731 5.53836 21.1673 5.53836 21.7846H4.43066C4.43066 20.9538 4.55182 20.1606 4.79413 19.4048C5.03643 18.6433 5.3797 17.9452 5.82393 17.3106C6.26816 16.6759 6.8047 16.1163 7.43355 15.6317C8.06816 15.1471 8.77489 14.7634 9.55374 14.4808C9.10374 14.2384 8.69989 13.9442 8.3422 13.5981C7.98451 13.2519 7.67874 12.8683 7.42489 12.4471C7.17682 12.0202 6.98355 11.5673 6.84509 11.0884C6.71239 10.6038 6.64605 10.1077 6.64605 9.59999C6.64605 8.83268 6.79028 8.11441 7.07874 7.44518C7.3672 6.77018 7.76239 6.18172 8.26432 5.67979C8.76624 5.17787 9.35182 4.78268 10.021 4.49422C10.696 4.20575 11.4172 4.06152 12.1845 4.06152C12.9518 4.06152 13.6701 4.20575 14.3393 4.49422C15.0143 4.78268 15.6028 5.17787 16.1047 5.67979C16.6066 6.18172 17.0018 6.77018 17.2903 7.44518C17.5787 8.11441 17.723 8.83268 17.723 9.59999C17.723 10.1077 17.6537 10.6009 17.5153 11.0798C17.3826 11.5586 17.1893 12.0086 16.9355 12.4298C16.6874 12.8509 16.3845 13.2375 16.0268 13.5894C15.6749 13.9356 15.2768 14.2298 14.8326 14.4721ZM7.75374 9.59999C7.75374 10.2115 7.86913 10.7856 8.0999 11.3221C8.33643 11.8586 8.65374 12.3288 9.05182 12.7327C9.45566 13.1308 9.92586 13.4481 10.4624 13.6846C10.9989 13.9154 11.573 14.0308 12.1845 14.0308C12.796 14.0308 13.3701 13.9154 13.9066 13.6846C14.4432 13.4481 14.9105 13.1308 15.3085 12.7327C15.7124 12.3288 16.0297 11.8586 16.2605 11.3221C16.497 10.7856 16.6153 10.2115 16.6153 9.59999C16.6153 8.98845 16.497 8.41441 16.2605 7.87787C16.0297 7.34133 15.7124 6.87402 15.3085 6.47595C14.9105 6.0721 14.4432 5.75479 13.9066 5.52402C13.3701 5.28748 12.796 5.16922 12.1845 5.16922C11.573 5.16922 10.9989 5.28748 10.4624 5.52402C9.92586 5.75479 9.45566 6.0721 9.05182 6.47595C8.65374 6.87402 8.33643 7.34133 8.0999 7.87787C7.86913 8.41441 7.75374 8.98845 7.75374 9.59999Z" fill="#E0E0E0"/>
    </g>
    </svg>
  }

}
