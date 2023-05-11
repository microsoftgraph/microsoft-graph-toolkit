/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client';
import { Configuration, InteractionRequiredAuthError, SilentRequest } from '@azure/msal-browser';
import {
  GraphEndpoint,
  loginContext,
  LoginType,
  MICROSOFT_GRAPH_DEFAULT_ENDPOINT,
  ProviderState,
  TeamsHelper,
  TeamsLib
} from '@microsoft/mgt-element';
import { Msal2Provider, PromptType } from '@microsoft/mgt-msal2-provider';

// eslint-disable-next-line @typescript-eslint/tslint/config
declare global {
  // eslint-disable-next-line @typescript-eslint/tslint/config
  interface Window {
    // eslint-disable-next-line @typescript-eslint/tslint/config
    nativeInterface: any;
  }
}

// eslint-disable-next-line @typescript-eslint/tslint/config
interface TeamsAuthResponse {
  // eslint-disable-next-line @typescript-eslint/tslint/config
  error?: string;
  // eslint-disable-next-line @typescript-eslint/tslint/config
  access_token?: string;
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
  GET = 'GET',

  /**
   * Will use Post and the request body
   */
  POST = 'POST'
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
   * @memberof AuthParams
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
 * Interface to define the configuration when creating a TeamsMsal2Provider
 *
 * @export
 * @interface TeamsMsal2Config
 */
export interface TeamsMsal2Config {
  /**
   * The app clientId
   *
   * @type {string}
   * @memberof TeamsMsal2Config
   */
  clientId: string;
  /**
   * The relative or absolute path of the html page that will handle the authentication
   *
   * @type {string}
   * @memberof TeamsMsal2Config
   */
  authPopupUrl: string;
  /**
   * The scopes to use when authenticating the user
   *
   * @type {string[]}
   * @memberof TeamsMsal2Config
   */
  scopes?: string[];
  /**
   * Additional Msal configurations options to use
   * See Msal.js documentation for more details
   *
   * @type {Configuration}
   * @memberof TeamsMsal2Config
   */
  msalOptions?: Configuration;
  /**
   * The relative or absolute path to the token exchange backend service
   *
   * @type {string}
   * @memberof TeamsMsal2Config
   */
  ssoUrl?: string;
  /**
   * Should the provider display a consent popup automatically if needed
   *
   * @type {string}
   * @memberof TeamsMsal2Config
   */
  autoConsent?: boolean;
  /**
   * Should the provider display a consent popup automatically if needed
   *
   * @type {AuthMethod}
   * @memberof TeamsMsal2Config
   */
  httpMethod?: HttpMethod;
  /**
   * The base URL for the graph client
   */
  baseURL?: GraphEndpoint;
}

/**
 * Enables authentication of Single page apps inside of a Microsoft Teams tab
 *
 * @deprecated in favor of TeamsFxProvider.
 * @export
 * @class TeamsMsal2Provider
 * @extends {Msal2Provider}
 */
export class TeamsMsal2Provider extends Msal2Provider {
  /**
   * Gets whether the Teams provider can be used in the current context
   * (Whether the app is running in Microsoft Teams)
   *
   * @readonly
   * @static
   * @memberof TeamsMsal2Provider
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
   * @memberof TeamsMsal2Provider
   */
  public static get microsoftTeamsLib(): TeamsLib {
    return TeamsHelper.microsoftTeamsLib;
  }
  public static set microsoftTeamsLib(value: TeamsLib) {
    TeamsHelper.microsoftTeamsLib = value;
  }

  /**
   * Name used for analytics
   *
   * @readonly
   * @memberof IProvider
   */
  public get name() {
    return 'MgtTeamsMsal2Provider';
  }

  /**
   * Handle all authentication redirects in the authentication page and authenticates the user
   *
   * @static
   * @returns
   * @memberof TeamsMsal2Provider
   */
  public static async handleAuth(baseURL: GraphEndpoint = MICROSOFT_GRAPH_DEFAULT_ENDPOINT) {
    // we are in popup world now - authenticate and handle it
    const teams = TeamsHelper.microsoftTeamsLib;
    if (!teams) {
      // eslint-disable-next-line no-console
      console.error('Make sure you have referenced the Microsoft Teams sdk before using the TeamsMsal2Provider');
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
      authParams = JSON.parse(paramsString) as AuthParams;
    } else {
      authParams = {};
    }

    if (!authParams.clientId) {
      teams.authentication.notifyFailure('no clientId provided');
      return;
    }

    const scopes = authParams.scopes ? authParams.scopes.split(',') : null;
    const prompt = authParams.isConsent ? PromptType.CONSENT : PromptType.SELECT_ACCOUNT;

    const options = authParams.options || { auth: { clientId: authParams.clientId } };

    options.system = options.system || {};
    options.system.loadFrameTimeout = 10000;

    const provider = new Msal2Provider({
      clientId: authParams.clientId,
      options,
      scopes,
      loginHint: authParams.loginHint,
      prompt,
      baseURL
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
          const isInIframe = window.parent !== window;
          if (!isInIframe) {
            sessionStorage.setItem(this._sessionStorageLoginInProgress, 'true');
          } else {
            // eslint-disable-next-line no-console
            console.warn(
              'handleProviderState - Is in iframe... will try to login anyway... but will not set session storage variable'
            );
          }

          await provider.login();
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
          teams.authentication.notifyFailure(e as string);
        }
      }
    };

    const handleProviderStateSync = () => {
      void handleProviderState();
    };

    provider.onStateChanged(handleProviderStateSync);
    await handleProviderState();
  }

  // eslint-disable-next-line @typescript-eslint/tslint/config
  protected clientId: string;

  private static _localStorageParametersKey = 'msg-TeamsMsal2Provider-auth-parameters';
  private static _sessionStorageLoginInProgress = 'msg-TeamsMsal2Provider-login-in-progress';
  private static _sessionStorageLogoutInProgress = 'msg-TeamsMsal2Provider-logout-in-progress';

  private teamsContext: loginContext;
  private _authPopupUrl: string;
  private _msalOptions: Configuration;
  private _ssoUrl: string;
  private _needsConsent: boolean;
  private _autoConsent: boolean;
  private _httpMethod: HttpMethod;

  constructor(config: TeamsMsal2Config) {
    super({
      clientId: config.clientId,
      loginType: LoginType.Redirect,
      options: config.msalOptions,
      scopes: config.scopes,
      baseURL: config.baseURL
    });
    this.clientId = config.clientId;
    this._msalOptions = config.msalOptions;
    this._authPopupUrl = config.authPopupUrl;
    this._ssoUrl = config.ssoUrl;
    this._autoConsent = typeof config.autoConsent !== 'undefined' ? config.autoConsent : true;
    this._httpMethod = typeof config.httpMethod !== 'undefined' ? config.httpMethod : HttpMethod.GET;

    const teams = TeamsHelper.microsoftTeamsLib;
    teams.initialize();

    // If we are in SSO-mode.
    if (this._ssoUrl) {
      void this.internalLogin();
    }
  }

  /**
   * Opens the teams authentication popup to the authentication page
   *
   * @returns {Promise<void>}
   * @memberof TeamsMsal2Provider
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
      void teams.getContext(context => {
        this.teamsContext = context;

        const loginHint = context.loginHint;
        const authParams: Partial<AuthParams> = {
          clientId: this.clientId,
          loginHint,
          options: this._msalOptions,
          scopes: this.scopes.join(','),
          isConsent: this._autoConsent
        };

        localStorage.setItem(TeamsMsal2Provider._localStorageParametersKey, JSON.stringify(authParams));

        const url = new URL(this._authPopupUrl, new URL(window.location.href));

        teams.authentication.authenticate({
          // eslint-disable-next-line @typescript-eslint/no-unused-vars
          failureCallback: _reason => {
            this.setState(ProviderState.SignedOut);
            reject();
          },
          // eslint-disable-next-line @typescript-eslint/no-unused-vars
          successCallback: _result => {
            // If we are in SSO Mode, the consent has been successful. Consider logged in
            if (this._ssoUrl) {
              this.setState(ProviderState.SignedIn);
              resolve();
            } else {
              // Otherwise log in via MSAL
              void this.trySilentSignIn();
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
      void teams.getContext(context => {
        this.teamsContext = context;

        const url = new URL(this._authPopupUrl, new URL(window.location.href));
        url.searchParams.append('signout', 'true');

        teams.authentication.authenticate({
          // eslint-disable-next-line @typescript-eslint/no-unused-vars
          failureCallback: _reason => {
            void this.trySilentSignIn();
            reject();
          },
          // eslint-disable-next-line @typescript-eslint/no-unused-vars
          successCallback: _result => {
            void this.trySilentSignIn();
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
   * @memberof TeamsMsal2Provider
   */
  public async getAccessToken(options: AuthenticationProviderOptions): Promise<string> {
    if (!this.teamsContext && TeamsHelper.microsoftTeamsLib) {
      const teams = TeamsHelper.microsoftTeamsLib;
      teams.initialize();
      try {
        this.teamsContext = await this.getTeamsContext();
      } catch (e) {
        throw new Error('Could not get teams context');
      }
    }

    const scopes = options ? options.scopes || this.scopes : this.scopes;

    // If we are in SSO Mode
    if (this._ssoUrl) {
      // Get token via the Teams SDK
      const clientToken = await this.getClientToken();

      const url: URL = new URL(this._ssoUrl, new URL(window.location.href));
      let response: Response;

      // Use GET and Query String
      if (this._httpMethod === HttpMethod.GET) {
        const params = new URLSearchParams({
          ssoToken: clientToken,
          scopes: scopes.join(','),
          clientId: this.clientId
        });

        response = await fetch(`${url.href}?${params.toString()}`, {
          method: 'GET',
          mode: 'cors',
          cache: 'default'
        });
      }
      // Use POST and body
      else {
        response = await fetch(url.href, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            authorization: `Bearer ${clientToken}`
          },
          body: JSON.stringify({
            scopes,
            clientid: this.clientId
          }),
          mode: 'cors',
          cache: 'default'
        });
      }
      try {
        // Exchange token from server
        const data: TeamsAuthResponse = (await response.json()) as TeamsAuthResponse;

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
      } catch (e) {
        this.unhandledFetchError(e);
      }
    }
    // If we are not in SSO Mode and using the Login component
    else {
      return new Promise((resolve, reject) => {
        const accessTokenRequest: SilentRequest = {
          scopes,
          account: this.getAccount()
        };

        this.publicClientApplication
          .acquireTokenSilent(accessTokenRequest)
          .then(response => {
            resolve(response.accessToken);
          })
          .catch(e => {
            if (e instanceof InteractionRequiredAuthError) {
              // nothing we can do now until we can do incremental consent
              // return null;
              resolve(null);
            } else {
              // throw e;
              reject(e);
            }
          });
      });
    }
  }
  /**
   * Makes sure we can get an access token before considered logged in
   *
   * @returns {Promise<void>}
   * @memberof TeamsMsal2Provider
   */
  private async internalLogin(): Promise<void> {
    // Try to get access token
    const accessToken: string = await this.getAccessToken(null);

    // If we have an access token. Consider the user signed in
    if (accessToken) {
      this.setState(ProviderState.SignedIn);
    } else {
      // If we need to consent to additional scopes
      if (this._needsConsent) {
        // If autoconsent is configured. Display a popup where the user can consent
        if (this._autoConsent) {
          // We need to pass the scopes from the client side
          if (!this.scopes) {
            throw new Error('For auto consent, scopes must be provided');
          } else {
            await this.login();
            return;
          }
        } else {
          throw new Error('Auto consent is not configured. You need to consent to additional scopes');
        }
      }
    }
  }

  /**
   * Get a token via the Teams SDK
   *
   * @returns {Promise<string>}
   * @memberof TeamsMsal2Provider
   */
  private async getClientToken(): Promise<string> {
    const teams = TeamsHelper.microsoftTeamsLib;
    return new Promise((resolve, reject) => {
      teams.authentication.getAuthToken({
        successCallback: (result: string) => {
          resolve(result);
        },
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        failureCallback: _reason => {
          this.setState(ProviderState.SignedOut);
          reject();
        }
      });
    });
  }

  /**
   * Retrieves the Teams context
   */
  private async getTeamsContext(): Promise<loginContext> {
    return new Promise(resolve => {
      const teams = TeamsHelper.microsoftTeamsLib;
      teams.initialize();
      void teams.getContext(context => {
        resolve(context);
        return;
      });
    });
  }

  private unhandledFetchError(err: unknown) {
    const message: string =
      typeof err === 'string' ? err : typeof err === 'object' ? err.toString() : JSON.stringify(err);
    // eslint-disable-next-line no-console
    console.error(`There was an error during the server side token exchange: ${message}`);
  }
}
