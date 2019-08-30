/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';
import { Configuration, UserAgentApplication } from 'msal';
import { LoginType, ProviderState } from './IProvider';
import { MsalProvider } from './MsalProvider';

// tslint:disable-next-line: completed-docs
declare var microsoftTeams: any;

// tslint:disable-next-line: completed-docs
declare global {
  // tslint:disable-next-line: completed-docs
  interface Window {
    // tslint:disable-next-line: completed-docs
    nativeInterface: any;
  }
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
}

/**
 * Interface to define the configuration when creating a TeamsProvider
 *
 * @export
 * @interface TeamsConfig
 */
export interface TeamsConfig {
  /**
   * The app clientId
   *
   * @type {string}
   * @memberof TeamsConfig
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
}

/**
 * Enables authentication of Single page apps inside of a Microsoft Teams tab
 *
 * @export
 * @class TeamsProvider
 * @extends {MsalProvider}
 */
export class TeamsProvider extends MsalProvider {
  private set accessToken(value: string) {
    this._accessToken = value;
    if (value) {
      sessionStorage.setItem(TeamsProvider._sessionStorageTokenKey, value);
      this.setState(ProviderState.SignedIn);
    } else {
      sessionStorage.removeItem(TeamsProvider._sessionStorageTokenKey);
      this.setState(ProviderState.SignedOut);
    }
  }

  private get accessToken() {
    return this._accessToken;
  }

  /**
   * Gets whether the Teams provider can be used in the current context
   * (Whether the app is running in Microsoft Teams)
   *
   * @readonly
   * @static
   * @memberof TeamsProvider
   */
  public static get isAvailable() {
    if (window.parent === window.self && window.nativeInterface) {
      // In Teams mobile client
      return true;
    } else if (window.name === 'embedded-page-container' || window.name === 'extension-tab-frame') {
      // In Teams web/desktop client
      return true;
    } else {
      return false;
    }
  }

  /**
   * Optional entry point to the teams library
   * If this value is not set, the provider will attempt to use
   * the microsoftTeams global variable.
   *
   * @static
   * @memberof TeamsProvider
   */
  public static microsoftTeamsLib;

  /**
   * Handle all authentication redirects in the authentication page and authenticates the user
   *
   * @static
   * @returns
   * @memberof TeamsProvider
   */
  public static async handleAuth() {
    // we are in popup world now - authenticate and handle it
    const teams = TeamsProvider.microsoftTeamsLib || microsoftTeams;
    if (!teams) {
      // tslint:disable-next-line: no-console
      console.error('Make sure you have referenced the Microsoft Teams sdk before using the TeamsProvider');
      return;
    }

    if (!this.isAvailable) {
      return;
    }

    teams.initialize();

    // msal checks for the window.opener.msal to check if this is a popup authentication
    // and gets a false positive since teams opens a popup for the authentication.
    // in reality, we are doing a redirect authentication and need to act as if this is the
    // window initiating the authentication
    if (window.opener) {
      window.opener.msal = null;
    }

    const url = new URL(window.location.href);

    const paramsString = sessionStorage.getItem(this._sessionStorageParametersKey);
    let authParams: AuthParams;

    if (paramsString) {
      authParams = JSON.parse(paramsString);
    } else {
      authParams = {};
    }

    if (!authParams.clientId) {
      authParams.clientId = url.searchParams.get('clientId');
      authParams.scopes = url.searchParams.get('scopes');
      authParams.loginHint = url.searchParams.get('loginHint');

      sessionStorage.setItem(this._sessionStorageParametersKey, JSON.stringify(authParams));
    }

    if (!authParams.clientId) {
      teams.authentication.notifyFailure('no clientId provided');
      return;
    }

    const scopes = authParams.scopes ? authParams.scopes.split(',') : null;

    const provider = new MsalProvider({
      clientId: authParams.clientId,
      options: {
        auth: {
          clientId: authParams.clientId,
          redirectUri: url.protocol + '//' + url.host + url.pathname
        },
        system: {
          loadFrameTimeout: 10000
        }
      },
      scopes
    });

    if ((UserAgentApplication.prototype as any).urlContainsHash(window.location.hash)) {
      // the page should redirect again
      return;
    }

    const handleProviderState = async () => {
      // how do we handle when user can't sign in
      // change to promise and return status
      if (provider.state === ProviderState.SignedOut) {
        provider.login({
          loginHint: authParams.loginHint,
          scopes: scopes || provider.scopes
        });
      } else if (provider.state === ProviderState.SignedIn) {
        try {
          const accessToken = await provider.getAccessTokenForScopes(...provider.scopes);
          sessionStorage.removeItem(this._sessionStorageParametersKey);
          teams.authentication.notifySuccess(accessToken);
        } catch (e) {
          sessionStorage.removeItem(this._sessionStorageParametersKey);
          teams.authentication.notifyFailure(e);
        }
      }
    };

    provider.onStateChanged(handleProviderState);
    handleProviderState();
  }

  private static _sessionStorageTokenKey = 'mgt-teamsprovider-accesstoken';
  private static _sessionStorageParametersKey = 'msg-teamsprovider-auth-parameters';

  /**
   * Scopes used for authentication
   *
   * @type {string[]}
   * @memberof TeamsProvider
   */
  public scopes: string[];
  private _authPopupUrl: string;
  private _accessToken: string;

  constructor(config: TeamsConfig) {
    super({
      clientId: config.clientId,
      loginType: LoginType.Redirect,
      options: config.msalOptions,
      scopes: config.scopes
    });

    const teams = TeamsProvider.microsoftTeamsLib || microsoftTeams;

    this._authPopupUrl = config.authPopupUrl;
    teams.initialize();
    this.accessToken = sessionStorage.getItem(TeamsProvider._sessionStorageTokenKey);
  }

  /**
   * Opens the teams authentication popup to the authentication page
   *
   * @returns {Promise<void>}
   * @memberof TeamsProvider
   */
  public async login(): Promise<void> {
    this.setState(ProviderState.Loading);
    const teams = TeamsProvider.microsoftTeamsLib || microsoftTeams;

    return new Promise((resolve, reject) => {
      teams.getContext(context => {
        const url = new URL(this._authPopupUrl, new URL(window.location.href));
        url.searchParams.append('clientId', this.clientId);

        if (context.loginHint) {
          url.searchParams.append('loginHint', context.loginHint);
        }

        if (this.scopes) {
          url.searchParams.append('scopes', this.scopes.join(','));
        }

        teams.authentication.authenticate({
          failureCallback: reason => {
            this.accessToken = null;
            reject();
          },
          successCallback: result => {
            this.accessToken = result;
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
    if (!this.accessToken) {
      throw null;
    } else {
      return this.accessToken;
    }
  }
}
