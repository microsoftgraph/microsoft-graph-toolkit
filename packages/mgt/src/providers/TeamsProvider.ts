/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';
import { AuthenticationParameters, Configuration, UserAgentApplication } from 'msal';
import { TeamsHelper } from '../utils/TeamsHelper';
import { LoginType, ProviderState } from '@microsoft/mgt-element';
import { MsalProvider } from './MsalProvider';

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

  /**
   * Additional Msal configurations options to use
   * See Msal.js documentation for more details
   *
   * @type {Configuration}
   * @memberof TeamsConfig
   */
  options?: Configuration;
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
  /**
   * Gets whether the Teams provider can be used in the current context
   * (Whether the app is running in Microsoft Teams)
   *
   * @readonly
   * @static
   * @memberof TeamsProvider
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
   * @memberof TeamsProvider
   */
  public static get microsoftTeamsLib(): any {
    return TeamsHelper.microsoftTeamsLib;
  }
  public static set microsoftTeamsLib(value: any) {
    TeamsHelper.microsoftTeamsLib = value;
  }

  /**
   * Handle all authentication redirects in the authentication page and authenticates the user
   *
   * @static
   * @returns
   * @memberof TeamsProvider
   */
  public static async handleAuth() {
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

    const options = authParams.options || { auth: { clientId: authParams.clientId } };

    options.system = options.system || {};
    options.system.loadFrameTimeout = 10000;

    const provider = new MsalProvider({
      clientId: authParams.clientId,
      options,
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
        if (isSignOut) {
          teams.authentication.notifySuccess();
          return;
        }

        // make sure we are calling login only once
        if (!sessionStorage.getItem(this._sessionStorageLoginInProgress)) {
          sessionStorage.setItem(this._sessionStorageLoginInProgress, 'true');
          provider.login({
            loginHint: authParams.loginHint,
            scopes: scopes || provider.scopes
          });
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

  private static _localStorageParametersKey = 'msg-teamsprovider-auth-parameters';
  private static _sessionStorageLoginInProgress = 'msg-teamsprovider-login-in-progress';
  private static _sessionStorageLogoutInProgress = 'msg-teamsprovider-logout-in-progress';

  private teamsContext;
  private _authPopupUrl: string;
  private _msalOptions: Configuration;

  constructor(config: TeamsConfig) {
    super({
      clientId: config.clientId,
      loginType: LoginType.Redirect,
      options: config.msalOptions,
      scopes: config.scopes
    });

    this._msalOptions = config.msalOptions;
    this._authPopupUrl = config.authPopupUrl;

    const teams = TeamsHelper.microsoftTeamsLib;
    teams.initialize();
  }

  /**
   * Opens the teams authentication popup to the authentication page
   *
   * @returns {Promise<void>}
   * @memberof TeamsProvider
   */
  public async login(): Promise<void> {
    this.setState(ProviderState.Loading);
    const teams = TeamsHelper.microsoftTeamsLib;

    return new Promise((resolve, reject) => {
      teams.getContext(context => {
        this.teamsContext = context;

        const authParams: AuthParams = {
          clientId: this.clientId,
          loginHint: context.loginHint,
          options: this._msalOptions,
          scopes: this.scopes.join(',')
        };

        localStorage.setItem(TeamsProvider._localStorageParametersKey, JSON.stringify(authParams));

        const url = new URL(this._authPopupUrl, new URL(window.location.href));

        teams.authentication.authenticate({
          failureCallback: reason => {
            this.setState(ProviderState.SignedOut);
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
   * sign out user
   *
   * @returns {Promise<void>}
   * @memberof MsalProvider
   */
  public async logout(): Promise<void> {
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

    const scopes = options ? options.scopes || this.scopes : this.scopes;
    const accessTokenRequest: AuthenticationParameters = {
      scopes
    };

    if (this.teamsContext && this.teamsContext.loginHint) {
      accessTokenRequest.loginHint = this.teamsContext.loginHint;
    }

    try {
      const response = await this._userAgentApplication.acquireTokenSilent(accessTokenRequest);
      return response.accessToken;
    } catch (e) {
      if (this.requiresInteraction(e)) {
        // nothing we can do now until we can do incremental consent
        return null;
      } else {
        throw e;
      }
    }
  }
}
