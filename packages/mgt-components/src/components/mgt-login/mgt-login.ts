/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { property } from 'lit/decorators.js';
import { classMap } from 'lit/directives/class-map.js';
import {
  Providers,
  ProviderState,
  MgtTemplatedComponent,
  IProviderAccount,
  mgtHtml,
  customElement
} from '@microsoft/mgt-element';

import { AvatarSize, IDynamicPerson, ViewType } from '../../graph/types';
import { MgtFlyout } from '../sub-components/mgt-flyout/mgt-flyout';
import { getUserWithPhoto } from '../../graph/graph.userWithPhoto';
import { MgtPerson } from '../mgt-person/mgt-person';
import { PersonViewType } from '../mgt-person/mgt-person-types';

import { getSvg, SvgIcon } from '../../utils/SvgHelper';

import { styles } from './mgt-login-css';
import { strings } from './strings';

import '../../styles/style-helper';

import { fluentListbox, fluentProgressRing, fluentButton, fluentCard } from '@fluentui/web-components';
import { registerFluentComponents } from '../../utils/FluentComponents';
registerFluentComponents(fluentListbox, fluentProgressRing, fluentButton, fluentCard);

/**
 *  loginViewType describes the enum strings that can be passed in to determine
 *  size of the mgt-login control.
 */
export type LoginViewType = 'avatar' | 'compact' | 'full';

// tslint:disable-next-line: completed-docs
type PersonViewConfig = {
  // tslint:disable-next-line: completed-docs
  view: ViewType;
  // tslint:disable-next-line: completed-docs
  avatarSize: AvatarSize;
};

/**
 * Web component button and flyout control to facilitate Microsoft identity platform authentication
 *
 * @export
 * @class MgtLogin
 * @extends {MgtBaseComponent}
 *
 * @fires {CustomEvent<undefined>} loginInitiated - Fired when login is initiated by the user
 * @fires {CustomEvent<undefined>} loginCompleted - Fired when login completes
 * @fires {CustomEvent<undefined>} loginFailed - Fired when login fails
 * @fires {CustomEvent<undefined>} logoutInitiated - Fired when logout is initiated by the user
 * @fires {CustomEvent<undefined>} logoutCompleted - Fired when logout completed
 *
 * @template signed-in-button-content (dataContext: {personDetails, personImage})
 * @template signed-out-button-content (dataContext: null)
 * @template flyout-commands (dataContext: {handleSignOut})
 * @template flyout-person-details (dataContext: {personDetails, personImage})
 *
 * @cssprop --font-size - {Length} Login font size
 * @cssprop --font-weight - {Length} Login font weight
 * @cssprop --margin - {String} Margin size
 * @cssprop --button-color - {Color} Login button font color
 * @cssprop --button-color--hover - {Color} Login button font hover color
 * @cssprop --button-background-color--hover - {Color} Login background hover color
 * @cssprop --popup-background-color - {Color} Popup background color
 * @cssprop --popup-color - {Color} Popup font color
 * @cssprop --popup-command-font-size - {Length} Popup command font size
 * @cssprop --popup-command-margin - {String} margins for the logout command in the popup
 * @cssprop --popup-padding - {String} padding applied inside the popup
 * @cssprop --profile-spacing - {String} margin applied to the active account inside the popup
 * @cssprop --profile-spacing-full - {String} margin applied to the active account inside the popup when login-view is full or more that one account is signed in.
 * @cssprop --add-account-button-color - {Color} Color for the text and icon of the add account button
 */
@customElement('login')
export class MgtLogin extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }
  /**
   * Returns the object of strings for localization
   *
   * @readonly
   * @protected
   * @memberof MgtLogin
   */
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
   * determines if presence is shown for logged in user
   * defaults to false
   * @type {boolean}
   */
  @property({
    attribute: 'show-presence',
    type: Boolean
  })
  public showPresence = false;

  /**
   * determines the view style to apply to the logged in user
   * options are 'full', 'compact', 'avatar', defaults to 'full'
   * @type {LoginViewType}
   */
  @property({
    attribute: 'login-view',
    type: String
  })
  public loginView: LoginViewType = 'full';

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

  /**
   * Suffix for user details key
   *
   * @private
   * @type {string}
   * @memberof MgtLogin
   */
  private _userDetailsKey: string = '-userDetails';

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
    const provider = Providers.globalProvider;
    if (!provider.isMultiAccountSupportedAndEnabled && (this.userDetails || !this.fireCustomEvent('loginInitiated'))) {
      return;
    }
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
    if (provider && provider.isMultiAccountSupportedAndEnabled) {
      localStorage.removeItem(provider.getActiveAccount().id + this._userDetailsKey);
    }
    if (provider && provider.logout) {
      await provider.logout();
      this.userDetails = null;
      if (provider.isMultiAccountSupportedAndEnabled) {
        localStorage.removeItem(provider.getActiveAccount().id + this._userDetailsKey);
      }
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
    const classes = {
      root: true,
      'vertical-layout': this.usesVerticalPersonCard
    };
    return html`
       <div class=${classMap(classes)} dir=${this.direction}>
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

        if (provider.isMultiAccountSupportedAndEnabled) {
          localStorage.setItem(
            Providers.globalProvider.getActiveAccount().id + this._userDetailsKey,
            JSON.stringify(this.userDetails)
          );
        }
        this.fireCustomEvent('loginCompleted');
      } else {
        this.userDetails = null;
      }
    }
  }

  private buildAriaLabel(isSignedIn: boolean, defaultLabel: string) {
    if (!isSignedIn) return defaultLabel;

    return (defaultLabel = this.userDetails ? this.userDetails.displayName : this.strings.signInLinkSubtitle);
  }

  /**
   * Render the button.
   *
   * @protected
   * @memberof MgtLogin
   */
  protected renderButton() {
    const isSignedIn = Providers.globalProvider?.state === ProviderState.SignedIn;
    const ariaLabel = this.buildAriaLabel(isSignedIn, this.strings.signInLinkSubtitle);

    const classes = {
      'login-button': true,
      'signed-in': isSignedIn,
      small: this.loginView === 'avatar',
      'full-size': this.loginView === 'full',
      'no-click': this._isFlyoutOpen
    };
    // uses a regular button for the signed in state to ease styling
    return isSignedIn && this.userDetails
      ? html`
        <button
          aria-label=${ariaLabel}
          @click=${this.onClick}
          class=${classMap(classes)}
        >
          ${this.renderSignedInButtonContent(this.userDetails, this._image)}
        </button>
      `
      : html`
        <fluent-button
          appearance="neutral"
          aria-label=${ariaLabel}
          ?disabled=${this.isLoadingState}
          @click=${this.onClick}
          class=${classMap(classes)}
        >
          ${this.renderSignedOutButtonContent()}
        </fluent-button>
      `;
  }

  private flyoutOpened = () => {
    this._isFlyoutOpen = true;
  };
  private flyoutClosed = () => {
    this._isFlyoutOpen = false;
  };

  /**
   * Render the details flyout.
   *
   * @protected
   * @memberof MgtLogin
   */
  protected renderFlyout() {
    return mgtHtml`
      <mgt-flyout
        class="flyout"
        light-dismiss
        @opened=${this.flyoutOpened}
        @closed=${this.flyoutClosed}
      >
        <div slot="flyout">
          <!-- Setting the card fill ensures the correct colors on hover states -->
          <fluent-card card-fill-color="#fbfbfb">
            ${this.renderFlyoutContent()}
          </fluent-card>
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
           <div class="popup-commands">
             ${this.renderFlyoutCommands()}
           </div>
           <div class="inside-content">
             <div class="main-profile">
               ${this.renderFlyoutPersonDetails(this.userDetails, this._image)}
             </div>
             ${this.renderAccounts()}
           </div>
           ${this.renderAddAccountContent()}
         </div>
       </div>
     `;
  }

  private get hasMultipleAccounts(): boolean {
    return (
      Providers.globalProvider?.isMultiAccountSupportedAndEnabled &&
      Providers.globalProvider?.getAllAccounts?.()?.length > 1
    );
  }

  private get usesVerticalPersonCard(): boolean {
    return this.loginView === 'full' || this.hasMultipleAccounts;
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
      mgtHtml`
        <mgt-person
          .personDetails=${personDetails}
          .personImage=${personImage}
          .view=${ViewType.twolines}
          .line2Property=${'email'}
          ?vertical-layout=${this.usesVerticalPersonCard}
          class="person"
        />
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
            <fluent-button
              appearance="lightweight"
              class="popup-command"
              @click=${this.logout}
              aria-label=${this.strings.signOutLinkSubtitle}
            >
              ${this.strings.signOutLinkSubtitle}
            </fluent-button>
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
   * Renders multi account content to add additional users
   *
   * @protected
   * @returns
   * @memberof MgtLogin
   */
  protected renderAddAccountContent() {
    if (Providers.globalProvider.isMultiAccountSupportedAndEnabled) {
      return html`
          <div class="add-account">
             <fluent-button
               appearance="lightweight"
               class="add-account-button"
               aria-label="Sign in with different account"
               @click=${() => {
                 this.login();
               }}
             >
               <i class="account-switch-icon">${getSvg(SvgIcon.SelectAccount, '#000000')}</i>
               Sign in with a different account
             </fluent-button>
           </div>
       `;
    }
  }

  private parsePersonDisplayConfiguration(): PersonViewConfig {
    const displayConfig: PersonViewConfig = { view: ViewType.twolines, avatarSize: 'small' };
    switch (this.loginView) {
      case 'avatar':
        displayConfig.view = ViewType.image;
        displayConfig.avatarSize = 'small';
        break;
      case 'compact':
        displayConfig.view = ViewType.oneline;
        displayConfig.avatarSize = 'small';
        break;
      case 'full':
      default:
        displayConfig.view = ViewType.twolines;
        displayConfig.avatarSize = 'large';
        break;
    }
    return displayConfig;
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
    const displayConfig = this.parsePersonDisplayConfiguration();
    return (
      template ||
      mgtHtml`
        <mgt-person
          .personDetails=${this.userDetails}
          .personImage=${this._image}
          .view=${displayConfig.view}
          .showPresence=${this.showPresence}
          .avatarSize=${displayConfig.avatarSize}
          line2-property="email"
          class="person"
        />
       `
    );
  }

  /**
   * POC for multi accounts - temporary
   *
   * @return {*}
   * @memberof MgtLogin
   */
  renderAccounts() {
    if (
      Providers.globalProvider.state === ProviderState.SignedIn &&
      Providers.globalProvider.isMultiAccountSupportedAndEnabled
    ) {
      const provider = Providers.globalProvider;
      const list = provider.getAllAccounts();

      if (list && list.length > 1) {
        return html`
         <div id="accounts">
           <fluent-design-system-provider>
             <fluent-listbox class="list-box" name="Account list">
              ${list.map(account => {
                if (account.id !== provider.getActiveAccount().id) {
                  const details = localStorage.getItem(account.id + this._userDetailsKey);
                  return mgtHtml`
                    <fluent-option class="list-box-option" value="${account.name}" role="option">
                      <mgt-person
                        @click=${() => this.setActiveAccount(account)}
                        @keyup=${(e: KeyboardEvent) => {
                          if (e.key === 'Enter') {
                            this.setActiveAccount(account);
                          }
                        }}
                        .personDetails=${details ? JSON.parse(details) : null}
                        .fallbackDetails=${{ displayName: account.name, mail: account.mail }}
                        .view=${PersonViewType.twolines}
                        class="person"
                      />
                    </fluent-option>
                  `;
                }
              })}
             </fluent-listbox>
           </fluent-design-system-provider>
         </div>
       `;
      }
    }
  }

  /**
   * Set one of the non-active accounts as the active account
   *
   * @param {*} account
   * @memberof MgtLogin
   */
  private setActiveAccount(account: IProviderAccount) {
    Providers.globalProvider.setActiveAccount(account);
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
    if (this.userDetails && this._isFlyoutOpen) {
      this.hideFlyout();
    } else if (this.userDetails) {
      this.showFlyout();
    } else {
      this.login();
    }
  }
}
