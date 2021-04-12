/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';
import { Configuration, InteractionRequiredAuthError } from '@azure/msal-browser';
import { LoginType, ProviderState, TeamsHelper } from '@microsoft/mgt-element';
// import { MsalProvider } from '@microsoft/mgt-msal2-provider';
import { Msal2Provider } from '@microsoft/mgt-msal2-provider';

// tslint:disable-next-line: completed-docs
declare global {
  // tslint:disable-next-line: completed-docs
  interface Window {
    // tslint:disable-next-line: completed-docs
    nativeInterface: any;
  }
}

/**
 * Enumeration to define if we are using get or post
 *
 * @export
 * @enum {string}
 */
export enum HttpMethod {
  /**
   * Will use Get and the querystring
   */
  get = 'GET',

  /**
   * Will use Post and the request body
   */
  post = 'POST'
}

/**
 * Interface used to store authentication parameters in session storage
 * between redirects
 *
 * @interface AuthParams
 */
interface AuthParams {
  /**
   * The app clientId
   *
   * @type {string}
   * @memberof AuthParams
   */
  clientId?: string;

  /**
   * The comma separated scopes
   *
   * @type {string}
   * @memberof AuthParams
   */
  scopes?: string;

  /**
   * The login hint to be used for authentication
   *
   * @type {string}
   * @memberof AuthParams
   */
  loginHint?: string;

  /**
   * Additional Msal configurations options to use
   * See Msal.js documentation for more details
   *
   * @type {Configuration}
   * @memberof TeamsConfig
   */
  options?: Configuration;
  /**
   * The login hint to be used for authentication
   *
   * @type {string}
   * @memberof AuthParams
   */
  isConsent?: boolean;
}

/**
 * Interface to define the configuration when creating a TeamsProvider
 *
 * @export
 * @interface Teams2Config
 */
export interface Teams2Config {
  /**
   * The app clientId
   *
   * @type {string}
   * @memberof Teams2Config
   */
  clientId: string;
  /**
   * The relative or absolute path of the html page that will handle the authentication
   *
   * @type {string}
   * @memberof TeamsConfig
   */
  authPopupUrl: string;
  /**
   * The scopes to use when authenticating the user
   *
   * @type {string[]}
   * @memberof TeamsConfig
   */
  scopes?: string[];
  /**
   * Additional Msal configurations options to use
   * See Msal.js documentation for more details
   *
   * @type {Configuration}
   * @memberof TeamsConfig
   */
  msalOptions?: Configuration;
  /**
   * The relative or absolute path to the token exchange backend service
   *
   * @type {string}
   * @memberof TeamsConfig
   */
  ssoUrl?: string;
  /**
   * Should the provider display a consent popup automatically if needed
   *
   * @type {string}
   * @memberof TeamsConfig
   */
  autoConsent?: boolean;
  /**
   * Should the provider display a consent popup automatically if needed
   *
   * @type {AuthMethod}
   * @memberof TeamsConfig
   */
  httpMethod?: HttpMethod;
}

/**
 * Enables authentication of Single page apps inside of a Microsoft Teams tab
 *
 * @export
 * @class Teams2Provider
 * @extends {Msal2Provider}
 */
export class Teams2Provider extends Msal2Provider {
  /**
   * Gets whether the Teams provider can be used in the current context
   * (Whether the app is running in Microsoft Teams)
   *
   * @readonly
   * @static
   * @memberof Teams2Provider
   */
  public static get isAvailable(): boolean {
    return TeamsHelper.isAvailable;
  }

  /**
   * Optional entry point to the teams library
   * If this value is not set, the provider will attempt to use
   * the microsoftTeams global variable.
   *
   * @static
   * @memberof Teams2Provider
   */
  public static get microsoftTeamsLib(): any {
    return TeamsHelper.microsoftTeamsLib;
  }
  public static set microsoftTeamsLib(value: any) {
    TeamsHelper.microsoftTeamsLib = value;
  }

  /**
   * Name used for analytics
   *
   * @readonly
   * @memberof IProvider
   */
  public get name() {
    return 'MgtTeams2Provider';
  }

  /**
   * Handle all authentication redirects in the authentication page and authenticates the user
   *
   * @static
   * @returns
   * @memberof Teams2Provider
   */
  public static async handleAuth() {
    debugger;
    // we are in popup world now - authenticate and handle it
    const teams = TeamsHelper.microsoftTeamsLib;
    if (!teams) {
      // tslint:disable-next-line: no-console
      console.error('Make sure you have referenced the Microsoft Teams sdk before using the TeamsProvider');
      return;
    }

    teams.initialize();

    // if we were signing out before, then we are done
    if (sessionStorage.getItem(this._sessionStorageLogoutInProgress)) {
      teams.authentication.notifySuccess();
    }

    const url = new URL(window.location.href);
    const isSignOut = url.searchParams.get('signout');

    const paramsString = localStorage.getItem(this._localStorageParametersKey);
    let authParams: AuthParams;

    if (paramsString) {
      authParams = JSON.parse(paramsString);
    } else {
      authParams = {};
    }

    if (!authParams.clientId) {
      teams.authentication.notifyFailure('no clientId provided');
      return;
    }

    const scopes = authParams.scopes ? authParams.scopes.split(',') : null;
    const prompt = authParams.isConsent ? 'consent' : null;

    const options = authParams.options || { auth: { clientId: authParams.clientId } };

    options.system = options.system || {};
    options.system.loadFrameTimeout = 10000;

    const provider = new Msal2Provider({
      clientId: authParams.clientId,
      options,
      scopes,
      loginHint: authParams.loginHint,
      prompt: prompt
    });

    const handleProviderState = async () => {
      // how do we handle when user can't sign in
      // change to promise and return status
      if (provider.state === ProviderState.SignedOut) {
        if (isSignOut) {
          teams.authentication.notifySuccess();
          return;
        }

        // make sure we are calling login only once
        if (!sessionStorage.getItem(this._sessionStorageLoginInProgress)) {
          sessionStorage.setItem(this._sessionStorageLoginInProgress, 'true');

          provider.login();
        }
      } else if (provider.state === ProviderState.SignedIn) {
        if (isSignOut) {
          sessionStorage.setItem(this._sessionStorageLogoutInProgress, 'true');
          await provider.logout();
          return;
        }

        try {
          const accessToken = await provider.getAccessTokenForScopes(...provider.scopes);
          teams.authentication.notifySuccess(accessToken);
        } catch (e) {
          teams.authentication.notifyFailure(e);
        }
      }
    };

    provider.onStateChanged(handleProviderState);
    handleProviderState();
  }

  protected clientId: string;

  private static _localStorageParametersKey = 'msg-teams2provider-auth-parameters';
  private static _sessionStorageLoginInProgress = 'msg-teams2provider-login-in-progress';
  private static _sessionStorageLogoutInProgress = 'msg-teams2provider-logout-in-progress';

  private teamsContext;
  private _authPopupUrl: string;
  private _msalOptions: Configuration;
  private _ssoUrl: string;
  private _needsConsent: boolean;
  private _autoConsent: boolean;

  constructor(config: Teams2Config) {
    super({
      clientId: config.clientId,
      loginType: LoginType.Redirect,
      options: config.msalOptions,
      scopes: config.scopes
    });
    this.clientId = config.clientId;
    this._msalOptions = config.msalOptions;
    this._authPopupUrl = config.authPopupUrl;
    this._ssoUrl = config.ssoUrl;
    this._autoConsent = config.autoConsent || false;

    const teams = TeamsHelper.microsoftTeamsLib;
    teams.initialize();

    // SSO Mode
    if (this._ssoUrl) {
      this.internalLogin();
    }
  }

  /**
   * Opens the teams authentication popup to the authentication page
   *
   * @returns {Promise<void>}
   * @memberof Teams2Provider
   */
  public async login(): Promise<void> {
    // In SSO mode the login should not be able to be run via user click
    // this method is called from the SSO internal login process if we need to consent
    if (this._ssoUrl && !this._needsConsent) {
      return;
    }

    this.setState(ProviderState.Loading);
    const teams = TeamsHelper.microsoftTeamsLib;

    return new Promise((resolve, reject) => {
      teams.getContext(context => {
        this.teamsContext = context;

        const authParams: AuthParams = {
          clientId: this.clientId,
          loginHint: context.loginHint,
          options: this._msalOptions,
          scopes: this.scopes.join(','),
          isConsent: this._autoConsent
        };

        localStorage.setItem(Teams2Provider._localStorageParametersKey, JSON.stringify(authParams));

        const url = new URL(this._authPopupUrl, new URL(window.location.href));

        teams.authentication.authenticate({
          failureCallback: reason => {
            this.setState(ProviderState.SignedOut);
            reject();
          },
          successCallback: result => {
            // If we are in auto consent mode return the accesstoken after successful consent
            if (this._autoConsent) {
              this.setState(ProviderState.SignedIn);
              resolve();
            } else {
              // Otherwise log in
              this.trySilentSignIn();
              resolve();
            }
          },
          url: url.href
        });
      });
    });
  }

  /**
   * sign out user
   *
   * @returns {Promise<void>}
   * @memberof MsalProvider
   */
  public async logout(): Promise<void> {
    // In SSO mode the logout should not be able to be run at all
    if (this._ssoUrl) {
      return;
    }
    const teams = TeamsHelper.microsoftTeamsLib;

    return new Promise((resolve, reject) => {
      teams.getContext(context => {
        this.teamsContext = context;

        const url = new URL(this._authPopupUrl, new URL(window.location.href));
        url.searchParams.append('signout', 'true');

        teams.authentication.authenticate({
          failureCallback: reason => {
            this.trySilentSignIn();
            reject();
          },
          successCallback: result => {
            this.trySilentSignIn();
            resolve();
          },
          url: url.href
        });
      });
    });
  }

  /**
   * Returns an access token that can be used for making calls to the Microsoft Graph
   *
   * @param {AuthenticationProviderOptions} options
   * @returns {Promise<string>}
   * @memberof TeamsProvider
   */
  public async getAccessToken(options: AuthenticationProviderOptions): Promise<string> {
    if (!this.teamsContext && TeamsHelper.microsoftTeamsLib) {
      const teams = TeamsHelper.microsoftTeamsLib;
      teams.initialize();
      this.teamsContext = await teams.getContext();
    }

    debugger;

    const scopes = options ? options.scopes || this.scopes : this.scopes;

    // SSO Mode
    if (this._ssoUrl) {
      const url = new URL(this._ssoUrl, new URL(window.location.href));
      // Get token via SSO
      const clientToken = await this.getClientToken();

      const params = new URLSearchParams({
        ssoToken: clientToken,
        scopes: scopes.join(','),
        clientId: this.clientId
      });

      // Exchange token from server
      const response: Response = await fetch(`${url.href}?${params}`, {
        method: 'GET',
        mode: 'cors',
        cache: 'default'
      });
      const data = await response.json().catch(this.unhandledFetchError);

      if (!response.ok && data.error === 'consent_required') {
        // A consent_required error means the user must consent to the requested scope, or use MFA
        // If we are in the log in process, display a dialog
        this._needsConsent = true;
      } else if (!response.ok) {
        throw data;
      } else {
        this._needsConsent = false;
        return data.access_token;
      }
    } else {
      const accessTokenRequest = {
        scopes: scopes,
        loginHint: this.teamsContext.loginHint
      };

      // if (this.teamsContext && this.teamsContext.loginHint) {
      //   accessTokenRequest.loginHint = this.teamsContext.loginHint;
      // }

      try {
        const response = await this.publicClientApplication.acquireTokenSilent(accessTokenRequest);
        return response.accessToken;
      } catch (e) {
        if (e instanceof InteractionRequiredAuthError) {
          // nothing we can do now until we can do incremental consent
          return null;
        } else {
          throw e;
        }
      }
    }
  }
  /**
   * Makes sure we can get an access token before considered logged in
   *
   * @returns {Promise<void>}
   * @memberof TeamsSSOProvider
   */
  private async internalLogin(): Promise<void> {
    // Try to get access token
    const accessToken: string = await this.getAccessToken(null);

    // If we need to consent, auto consent mode is on, and we have scopes on the client side
    if (!accessToken && this._needsConsent && this._autoConsent && this.scopes) {
      this.login();
      return;
    }

    this.setState(accessToken ? ProviderState.SignedIn : ProviderState.SignedOut);
  }

  /**
   * Get a token via the Teams SDK
   *
   * @returns {Promise<string>}
   * @memberof TeamsSSOProvider
   */
  private async getClientToken(): Promise<string> {
    const teams = TeamsHelper.microsoftTeamsLib;
    return new Promise((resolve, reject) => {
      teams.authentication.getAuthToken({
        successCallback: (result: string) => {
          resolve(result);
        },
        failureCallback: reason => {
          this.setState(ProviderState.SignedOut);
          reject();
        }
      });
    });
  }

  private unhandledFetchError(err: any) {
    console.error(`There was an error during the server side token exchange: ${err}`);
  }
}
