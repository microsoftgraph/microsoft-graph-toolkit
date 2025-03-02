/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { CSSResult, html, TemplateResult } from 'lit';
import { property, state } from 'lit/decorators.js';
import { classMap } from 'lit/directives/class-map.js';
import { ifDefined } from 'lit/directives/if-defined.js';
import { Providers, ProviderState, IProviderAccount, mgtHtml, MgtTemplatedTaskComponent } from '@microsoft/mgt-element';

import { AvatarSize, IDynamicPerson, ViewType } from '../../graph/types';
import { MgtFlyout, registerMgtFlyoutComponent } from '../sub-components/mgt-flyout/mgt-flyout';
import { getUserWithPhoto } from '../../graph/graph.userWithPhoto';
import { MgtPerson, registerMgtPersonComponent } from '../mgt-person/mgt-person';

import { getSvg, SvgIcon } from '../../utils/SvgHelper';

import { styles } from './mgt-login-css';
import { strings } from './strings';

import '../../styles/style-helper';

import { fluentListbox, fluentProgressRing, fluentButton, fluentCard } from '@fluentui/web-components';
import { registerFluentComponents } from '../../utils/FluentComponents';
import { registerComponent } from '@microsoft/mgt-element';
import { TaskStatus } from '@lit/task';

/**
 * loginViewType describes the enum strings that can be passed in to determine
 * size of the mgt-login control.
 */
export type LoginViewType = 'avatar' | 'compact' | 'full';

interface PersonViewConfig {
  view: ViewType;
  avatarSize: AvatarSize;
}

export const registerMgtLoginComponent = () => {
  registerFluentComponents(fluentListbox, fluentProgressRing, fluentButton, fluentCard);

  registerMgtFlyoutComponent();
  registerMgtPersonComponent();
  registerComponent('login', MgtLogin);
};

/**
 * Web component button and flyout control to facilitate Microsoft identity platform authentication
 *
 * @export
 * @class MgtLogin
 * @extends {MgtTemplatedTaskComponent}
 *
 * @fires {CustomEvent<undefined>} updated - Fired when the component is updated
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
 * @cssprop --login-signed-in-background - {String} the background properties of the component when signed in.
 * @cssprop --login-signed-in-hover-background - {String} the background properties of the component when signed in.
 * @cssprop --login-signed-out-button-background - {String} the background properties of the component when signed out.
 * @cssprop --login-signed-out-button-hover-background - {String} the background properties of the component when signed out.
 * @cssprop --login-signed-out-button-text-color - {Color} the background color of the component when signed out.
 * @cssprop --login-button-padding - {Length} the padding of the button. Default is 0px.
 * @cssprop --login-popup-background-color - {Color} the background color of the popup.
 * @cssprop --login-popup-command-button-background-color - {Color} the color of the background to the popup command button.
 * @cssprop --login-popup-padding - {Length} the padding applied to the popup card. Default is 16px.
 * @cssprop --login-add-account-button-text-color - {Color} the color for the text and icon of the add account button.
 * @cssprop --login-add-account-button-background-color - {Color} the color for the background and icon of the add account button.
 * @cssprop --login-add-account-button-hover-background-color - {Color} the color for the background and icon of the add account button on hover.
 * @cssprop --login-command-button-text-color - {Color} the color for the text of the command button.
 * @cssprop --login-command-button-background-color - {Color} the color for the background of the command button.
 * @cssprop --login-command-button-hover-background-color - {Color} the color for the background of the command button on hovering.
 * @cssprop --login-account-item-hover-bg-color - {Color} the background color of the account item on hover.
 * @cssprop --login-flyout-command-text-color - {Color} the color for the text of the flyout command button.
 * @cssprop --login-person-avatar-size - {Length} the size of the avatar in the person component. Default is 40px.
 */
export class MgtLogin extends MgtTemplatedTaskComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles(): CSSResult[] {
    return styles;
  }
  /**
   * Returns the object of strings for localization
   *
   * @readonly
   * @protected
   * @memberof MgtLogin
   */
  protected get strings(): Record<string, string> {
    return strings;
  }

  /**
   * Allows developer to use specific user details for login.
   *
   * @type {IDynamicPerson}
   */
  @property({
    attribute: 'user-details',
    type: Object
  })
  public userDetails: IDynamicPerson;

  /**
   * Determines if presence is shown for logged in user
   * defaults to false
   *
   * @type {boolean}
   */
  @property({
    attribute: 'show-presence',
    type: Boolean
  })
  public showPresence = false;

  /**
   * Determines the view style to apply to the logged in user
   * options are 'full', 'compact', 'avatar', defaults to 'full'
   *
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
   * Determines if login menu popup should be showing.
   *
   * @private
   * @type {boolean}
   */
  @state() private _isFlyoutOpen: boolean;

  /**
   * The image blob string
   *
   * @private
   * @type {string}
   * @memberof MgtLogin
   */
  private _image: string;

  /**
   * Suffix for user details key
   *
   * @private
   * @type {string}
   * @memberof MgtLogin
   */
  private get _userDetailsKey() {
    return '-userDetails';
  }

  @state() private _arrowKeyLocation = -1;

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
    // (If we have user details or the consumer doesn't cancel the loginInitiated event) and the provider doesn't support multi-account, we don't have to login.
    // This condition is to prevent the login popup from showing up when the user is already logged in while still ensuring the loginIntiated event is raised
    if ((this.userDetails || !this.fireCustomEvent('loginInitiated')) && !provider.isMultiAccountSupportedAndEnabled) {
      return;
    }
    if (provider?.login) {
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
  public logout = async (): Promise<void> => {
    if (!this.fireCustomEvent('logoutInitiated')) {
      return;
    }

    const provider = Providers.globalProvider;
    if (provider?.logout) {
      await provider.logout();
    }
  };

  private readonly completeLogout = () => {
    const provider = Providers.globalProvider;
    if (provider.state === ProviderState.SignedOut) {
      this.userDetails = null;
      if (provider.isMultiAccountSupportedAndEnabled) {
        const activeAccount = provider.getActiveAccount();
        localStorage.removeItem(activeAccount?.id + this._userDetailsKey);
      }
      this.hideFlyout();
      this.fireCustomEvent('logoutCompleted');
    }
  };

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   *
   * @protected
   * @returns {TemplateResult}
   */
  protected renderContent = (): TemplateResult => {
    return html`
      <div class="login-root">
        ${this.renderButton()}
        ${this.renderFlyout()}
      </div>`;
  };

  protected args(): unknown[] {
    return [this.providerState];
  }

  /**
   * Load state into the component.
   *
   * @protected
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
        if (provider.logout) {
          this.completeLogout();
        }
        this.userDetails = null;
      }
    }
  }

  /**
   * Render the sign in or sign out button.
   *
   * @protected
   * @memberof MgtLogin
   * @returns {TemplateResult}
   */
  protected renderButton(): TemplateResult {
    const isSignedIn = Providers.globalProvider?.state === ProviderState.SignedIn;
    const loginClasses = classMap({
      'signed-in': isSignedIn && Boolean(this.userDetails),
      'signed-out': !isSignedIn,
      small: this.loginView === 'avatar'
    });
    const appearance = isSignedIn ? 'stealth' : 'neutral';
    const showSignedInState = isSignedIn && this.userDetails;
    const buttonContentTemplate = showSignedInState
      ? this.renderSignedInButtonContent(this.userDetails, this._image)
      : this.renderSignedOutButtonContent();
    const expandedState: boolean | undefined = showSignedInState ? this._isFlyoutOpen : undefined;
    return html`
      <fluent-button
        id="login-button"
        aria-expanded="${ifDefined(expandedState)}"
        appearance=${appearance}
        aria-label="${ifDefined(isSignedIn ? undefined : this.strings.signInLinkSubtitle)}"
        ?disabled=${this._task.status === TaskStatus.PENDING}
        @click=${this.onClick}
        class=${loginClasses}>
          ${buttonContentTemplate}
      </fluent-button>`;
  }

  private readonly flyoutOpened = () => {
    this._isFlyoutOpen = true;
  };
  private readonly flyoutClosed = () => {
    this._isFlyoutOpen = false;
  };

  /**
   * Render the details flyout.
   *
   * @protected
   * @memberof MgtLogin
   * @returns {TemplateResult}
   */
  protected renderFlyout(): TemplateResult {
    return mgtHtml`
      <mgt-flyout
        class="flyout"
        light-dismiss
        @opened=${this.flyoutOpened}
        @closed=${this.flyoutClosed}>
        <fluent-card
          slot="flyout"
          tabindex="0"
          class="flyout-card"
          @keydown=${this.onUserKeyDown}
          >
          ${this.renderFlyoutContent()}
        </fluent-card>
      </mgt-flyout>`;
  }

  /**
   * Tracks tabbing through the flyout (keydown)
   */
  private readonly onUserKeyDown = (e: KeyboardEvent): void => {
    if (!this.flyout.isOpen) {
      return;
    }

    const el = this.renderRoot.querySelector('.popup-content');
    const focusableEls = el.querySelectorAll('ul, fluent-button');
    const firstFocusableEl = el.querySelector('#signout-button') || focusableEls[0];
    const lastFocusableEl =
      el.querySelector('#signin-different-account-button') || focusableEls[focusableEls.length - 1];

    if (e.key === 'Tab' && e.shiftKey && firstFocusableEl === e.target) {
      e.preventDefault();
      (lastFocusableEl as HTMLElement)?.focus();
    }
    if (e.key === 'Tab' && !e.shiftKey && lastFocusableEl === e.target) {
      e.preventDefault();
      (firstFocusableEl as HTMLElement)?.focus();
    }
    if (e.key === 'Escape') {
      const loginButton = this.renderRoot.querySelector('#login-button');
      (loginButton as HTMLElement)?.focus();
    }
    const fluentCardEl = this.renderRoot.querySelector('fluent-card');
    if (e.shiftKey && e.key === 'Tab' && e.target === fluentCardEl) {
      this.hideFlyout();
    }
  };

  /**
   * Render the flyout menu content.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtLogin
   */
  protected renderFlyoutContent(): TemplateResult {
    if (!this.userDetails) {
      return;
    }
    return html`
       <div class="popup">
         <div class="popup-content">
           <div class="commands">
             ${this.renderFlyoutCommands()}
           </div>
           <div class="content">
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
   * @returns {TemplateResult}
   * @memberof MgtLogin
   */
  protected renderFlyoutPersonDetails(personDetails: IDynamicPerson, personImage: string): TemplateResult {
    const template = this.renderTemplate('flyout-person-details', { personDetails, personImage });
    return (
      template ||
      mgtHtml`
        <mgt-person
          .personDetails=${personDetails}
          .personImage=${personImage}
          view="twolines"
          .line2Property=${'email'}
          ?vertical-layout=${this.usesVerticalPersonCard}
          class="person">
        </mgt-person>`
    );
  }

  /**
   * Render the flyout commands.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtLogin
   */
  protected renderFlyoutCommands(): TemplateResult {
    const template = this.renderTemplate('flyout-commands', { handleSignOut: () => this.logout() });
    return (
      template ||
      html`
        <fluent-button
          id="signout-button"
          appearance="stealth"
          size="medium"
          class="flyout-command"
          @click=${this.logout}
          aria-label=${this.strings.signOutLinkSubtitle}>
            ${this.strings.signOutLinkSubtitle}
        </fluent-button>`
    );
  }

  /**
   * Render the button content.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtLogin
   */
  protected renderButtonContent(): TemplateResult {
    if (this.userDetails) {
      return this.renderSignedInButtonContent(this.userDetails, this._image);
    } else {
      return this.renderSignedOutButtonContent();
    }
  }

  /**
   * Renders the button to allow adding accounts.
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
            id="signin-different-account-button"
            appearance="stealth"
            aria-label="${this.strings.signInWithADifferentAccount}"
            @click=${() => void this.login()}>
            <span slot="start"><i>${getSvg(SvgIcon.SelectAccount, 'currentColor')}</i></span>
            ${this.strings.signInWithADifferentAccount}
          </fluent-button>
        </div>`;
    }
  }

  private parsePersonDisplayConfiguration(): PersonViewConfig {
    const displayConfig: PersonViewConfig = { view: 'twolines', avatarSize: 'small' };
    switch (this.loginView) {
      case 'avatar':
        displayConfig.view = 'image';
        displayConfig.avatarSize = 'small';
        break;
      case 'compact':
        displayConfig.view = 'oneline';
        displayConfig.avatarSize = 'small';
        break;
      case 'full':
      default:
        displayConfig.view = 'twolines';
        displayConfig.avatarSize = 'auto';
        break;
    }
    return displayConfig;
  }

  /**
   * Render the button content when the user is signed in.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtLogin
   */
  protected renderSignedInButtonContent(personDetails: IDynamicPerson, personImage: string): TemplateResult {
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
          class="signed-in-person"
        ></mgt-person>`
    );
  }

  /**
   * Renders multiple accounts that can be used to sign in.
   *
   * @return {TemplateResult}
   * @memberof MgtLogin
   */
  renderAccounts(): TemplateResult {
    if (
      Providers.globalProvider.state === ProviderState.SignedIn &&
      Providers.globalProvider.isMultiAccountSupportedAndEnabled
    ) {
      const provider = Providers.globalProvider;
      const accounts = provider.getAllAccounts();

      if (accounts?.length > 1) {
        return html`
          <div id="accounts">
            <ul
              tabindex="0"
              class="account-list"
              part="account-list"
              aria-label="${this.ariaLabel}"
              @keydown=${this.handleAccountListKeyDown}
            >
              ${accounts
                .filter(a => a.id !== provider.getActiveAccount().id)
                .map(account => {
                  const details = localStorage.getItem(account.id + this._userDetailsKey);
                  return mgtHtml`
                    <li
                      tabindex="-1"
                      part="account-item"
                      class="account-item"
                      @click=${() => this.setActiveAccount(account)}
                      @keyup=${(e: KeyboardEvent) => {
                        if (e.key === 'Enter') this.setActiveAccount(account);
                      }}
                    >
                      <mgt-person
                        .personDetails=${details ? JSON.parse(details) : null}
                        .fallbackDetails=${{ displayName: account.name, mail: account.mail }}
                        .view=${'twolines'}
                        class="account"
                      ></mgt-person>
                    </li>`;
                })}
            </ul>
          </div>
       `;
      }
    }
  }

  private readonly handleAccountListKeyDown = (event: KeyboardEvent) => {
    const list: HTMLUListElement = this.renderRoot.querySelector('ul.account-list');
    let item: HTMLLIElement;
    const listItems: HTMLCollection = list?.children;
    // Default all tabindex values in li nodes to -1
    for (const element of listItems) {
      const el = element as HTMLLIElement;
      el.setAttribute('tabindex', '-1');
      el.blur();
    }

    const childElementCount = list.childElementCount;
    const keyName = event.key;
    if (keyName === 'ArrowDown') {
      this._arrowKeyLocation = (this._arrowKeyLocation + 1 + childElementCount) % childElementCount;
    } else if (keyName === 'ArrowUp') {
      this._arrowKeyLocation = (this._arrowKeyLocation - 1 + childElementCount) % childElementCount;
    } else if (keyName === 'Tab' || keyName === 'Escape') {
      this._arrowKeyLocation = -1;
      list.blur();
      if (keyName === 'Escape') {
        event.preventDefault();
        event.stopPropagation();
      }
      return;
    }

    if (this._arrowKeyLocation > -1) {
      item = listItems[this._arrowKeyLocation] as HTMLLIElement;
      item.setAttribute('tabindex', '1');
      item.focus();
    }
  };

  /**
   * Set one of the non-active accounts as the active account
   *
   * @param {IProviderAccount} account
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
  protected renderSignedOutButtonContent(): TemplateResult {
    const template = this.renderTemplate('signed-out-button-content', null);
    return (
      template ||
      html`
        <span>${this.strings.signInLinkSubtitle}</span>`
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

  /**
   * Handles the click on the button in the flyout.
   *
   * @private
   * @memberof MgtLogin
   */
  private readonly onClick = (): void => {
    if (this.userDetails && this.flyout.isOpen) {
      this.hideFlyout();
    } else if (this.userDetails) {
      this.showFlyout();
    } else {
      void this.login();
    }
  };
}
