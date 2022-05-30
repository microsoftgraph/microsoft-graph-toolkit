/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, property } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { Providers, ProviderState, MgtTemplatedComponent } from '@microsoft/mgt-element';

import { IDynamicPerson, ViewType } from '../../graph/types';
import { MgtFlyout } from '../sub-components/mgt-flyout/mgt-flyout';
import { getUserWithPhoto } from '../../graph/graph.userWithPhoto';
import { MgtPerson, PersonViewType } from '../mgt-person/mgt-person';

import { styles } from './mgt-login-css';
import { strings } from './strings';

import '../../styles/style-helper';

/**
 * Web component button and flyout control to facilitate Microsoft identity platform authentication
 *
 * @export
 * @class MgtLogin
 * @extends {MgtBaseComponent}
 *
 * @fires loginInitiated - Fired when login is initiated by the user
 * @fires loginCompleted - Fired when login completes
 * @fires loginFailed - Fired when login fails
 * @fires logoutInitiated - Fired when logout is initiated by the user
 * @fires logoutCompleted - Fired when logout completed
 *
 * @template signed-in-button-content (dataContext: {personDetails, personImage})
 * @template signed-out-button-content (dataContext: null)
 * @template flyout-commands (dataContext: {handleSignOut})
 * @template flyout-person-details (dataContext: {personDetails, personImage})
 *
 * @cssprop --font-size - {Length} Login font size
 * @cssprop --font-weight - {Length} Login font weight
 * @cssprop --height - {String} Login height percentage
 * @cssprop --margin - {String} Margin size
 * @cssprop --padding - {String} Padding size
 * @cssprop --button-color - {Color} Login button font color
 * @cssprop --button-color--hover - {Color} Login button font hover color
 * @cssprop --button-background-color - {Color} Login button background color
 * @cssprop --button-background-color--hover - {Color} Login background hover color
 * @cssprop --popup-background-color - {Color} Popup background color
 * @cssprop --popup-color - {Color} Popup font color
 * @cssprop --popup-command-font-size - {Length} Popup command font size
 */
@customElement('mgt-login')
export class MgtLogin extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }
  protected get strings() {
    return strings;
  }

  /**
   * allows developer to use specific user details for login
   * @type {IDynamicPerson}
   */
  @property({
    attribute: 'user-details',
    type: Object
  })
  public userDetails: IDynamicPerson;

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
   * Get the scopes required for login
   *
   * @static
   * @return {*}  {string[]}
   * @memberof MgtLogin
   */
  public static get requiredScopes(): string[] {
    return [...new Set(['user.read', ...MgtPerson.requiredScopes])];
  }

  /**
   * determines if login menu popup should be showing
   * @type {boolean}
   */
  @property({ attribute: false }) private _isFlyoutOpen: boolean;

  private _image: string;

  constructor() {
    super();
    this._isFlyoutOpen = false;
  }

  /**
   * Invoked each time the custom element is appended into a document-connected element
   *
   * @memberof MgtLogin
   */
  public connectedCallback() {
    super.connectedCallback();
    this.addEventListener('click', e => e.stopPropagation());
  }

  /**
   * Initiate login
   *
   * @returns {Promise<void>}
   * @memberof MgtLogin
   */
  public async login(): Promise<void> {
    if (this.userDetails || !this.fireCustomEvent('loginInitiated')) {
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
   * Initiate logout
   *
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
      this.hideFlyout();
      this.fireCustomEvent('logoutCompleted');
    }
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    return html`
      <div class="root" dir=${this.direction}>
        <div>
          ${this.renderButton()}
        </div>
        ${this.renderFlyout()}
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
    if (provider && !this.userDetails) {
      if (provider.state === ProviderState.SignedIn) {
        this.userDetails = await getUserWithPhoto(provider.graph.forComponent(this));

        if (this.userDetails.personImage) {
          this._image = this.userDetails.personImage;
        }

        this.fireCustomEvent('loginCompleted');
      } else {
        this.userDetails = null;
      }
    }
  }

  /**
   * Render the button.
   *
   * @protected
   * @memberof MgtLogin
   */
  protected renderButton() {
    const provider = Providers.globalProvider;
    const signedInState = ProviderState.SignedIn;

    let ariaLabel = this.strings.signInLinkSubtitle;
    if (provider) {
      if (provider.state === signedInState) {
        ariaLabel = this.userDetails ? this.userDetails.displayName : this.strings.signInLinkSubtitle;
      }
    }

    const classes = {
      'login-button': true,
      'no-click': this._isFlyoutOpen
    };
    return html`
      <button aria-label=${ariaLabel} ?disabled="${this.isLoadingState}" @click=${this.onClick} class=${classMap(
      classes
    )} role="button">
        ${this.renderButtonContent()}
      </button>
    `;
  }

  /**
   * Render the details flyout.
   *
   * @protected
   * @memberof MgtLogin
   */
  protected renderFlyout() {
    return html`
      <mgt-flyout
        class="flyout"
        light-dismiss
        @opened=${() => (this._isFlyoutOpen = true)}
        @closed=${() => (this._isFlyoutOpen = false)}
      >
        <div slot="flyout">
          ${this.renderFlyoutContent()}
        </div>
      </mgt-flyout>
    `;
  }

  /**
   * Render the flyout menu content.
   *
   * @protected
   * @returns
   * @memberof MgtLogin
   */
  protected renderFlyoutContent() {
    if (!this.userDetails) {
      return;
    }
    return html`
      <div class="popup">
        <div class="popup-content">
          <div>
            ${this.renderFlyoutPersonDetails(this.userDetails, this._image)}
          </div>
          <div class="popup-commands">
            ${this.renderFlyoutCommands()}
          </div>
        </div>
      </div>
    `;
  }

  /**
   * Render the flyout person details.
   *
   * @protected
   * @returns
   * @memberof MgtLogin
   */
  protected renderFlyoutPersonDetails(personDetails: IDynamicPerson, personImage: string) {
    const template = this.renderTemplate('flyout-person-details', { personDetails, personImage });
    return (
      template ||
      html`
        <mgt-person .personDetails=${personDetails} .personImage=${personImage} .view=${ViewType.twolines} />
      `
    );
  }

  /**
   * Render the flyout commands.
   *
   * @protected
   * @returns
   * @memberof MgtLogin
   */
  protected renderFlyoutCommands() {
    const template = this.renderTemplate('flyout-commands', { handleSignOut: () => this.logout() });
    return (
      template ||
      html`
        <ul>
          <li>
            <button class="popup-command" @click=${this.logout} aria-label=${this.strings.signOutLinkSubtitle}>
              ${this.strings.signOutLinkSubtitle}
            </button>
          </li>
        </ul>
      `
    );
  }

  /**
   * Render the button content.
   *
   * @protected
   * @returns
   * @memberof MgtLogin
   */
  protected renderButtonContent() {
    if (this.userDetails) {
      return this.renderSignedInButtonContent(this.userDetails, this._image);
    } else {
      return this.renderSignedOutButtonContent();
    }
  }

  /**
   * Render the button content when the user is signed in.
   *
   * @protected
   * @returns
   * @memberof MgtLogin
   */
  protected renderSignedInButtonContent(personDetails: IDynamicPerson, personImage: string) {
    const template = this.renderTemplate('signed-in-button-content', { personDetails, personImage });
    return (
      template ||
      html`
        <mgt-person .personDetails=${this.userDetails} .personImage=${this._image} .view=${ViewType.oneline} />
      `
    );
  }

  /**
   * Clears state of the component
   *
   * @protected
   * @memberof MgtLogin
   */
  protected clearState() {
    this.userDetails = null;
    this._image = null;
  }

  /**
   * Render the button content when the user is not signed in.
   *
   * @protected
   * @returns
   * @memberof MgtLogin
   */
  protected renderSignedOutButtonContent() {
    const template = this.renderTemplate('signed-out-button-content', null);
    return (
      template ||
      html`
        <i class="login-icon ms-Icon ms-Icon--Contact"></i>
        <span>${this.strings.signInLinkSubtitle}</span>
      `
    );
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

  private onClick() {
    if (this.userDetails) {
      this.showFlyout();
    } else {
      this.login();
    }
  }
}
