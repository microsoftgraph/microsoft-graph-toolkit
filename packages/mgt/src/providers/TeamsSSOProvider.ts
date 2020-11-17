/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';
import { TeamsHelper } from '../utils/TeamsHelper';
import { ProviderState, IProvider } from '@microsoft/mgt-element';
import { createFromProvider } from '../Graph';
import { UserAgentApplication } from 'msal';

// tslint:disable-next-line: completed-docs
declare global {
  // tslint:disable-next-line: completed-docs
  interface Window {
    // tslint:disable-next-line: completed-docs
    nativeInterface: any;
  }
}
/**
 * Interface to define the configuration when creating a TeamsSSOProvider
 *
 * @export
 * @interface TeamsSSOConfig
 */
export interface TeamsSSOConfig {
  /**
   * The app clientId
   *
   * @type {string}
   * @memberof TeamsConfig
   */
  clientId: string;
  /**
   * The relative or absolute path of the API that will exchange the token
   *
   * @type {string}
   * @memberof TeamsSSOConfig
   */
  ssoUrl: string;
  /**
   * The scopes that will be used in the application
   *
   * @type {string[]}
   * @memberof TeamsSSOConfig
   */
  scopes: string[];
}

/**
 * Enables authentication of Single page apps inside of a Microsoft Teams tab
 *
 * @export
 * @class TeamsSSOProvider
 * @extends {IProvider}
 */
export class TeamsSSOProvider extends IProvider {
  /**
   * Gets whether the Teams SSO provider can be used in the current context
   * (Whether the app is running in Microsoft Teams)
   *
   * @readonly
   * @static
   * @memberof TeamsSSOProvider
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
   * @memberof TeamsSSOProvider
   */
  public static get microsoftTeamsLib(): any {
    return TeamsHelper.microsoftTeamsLib;
  }
  public static set microsoftTeamsLib(value: any) {
    TeamsHelper.microsoftTeamsLib = value;
  }
  private _ssoConfig: TeamsSSOConfig;
  private _ssoUrl: string;
  private _needsConsent: boolean;

  constructor(config: TeamsSSOConfig) {
    super();
    this._ssoConfig = config;
    this._ssoUrl = config.ssoUrl;
    this._needsConsent = false;

    const teams = TeamsHelper.microsoftTeamsLib;
    teams.initialize();

    this.graph = createFromProvider(this);
    this.internalLogin();
  }

  /**
   * Opens a MSAL popup so we can consent to the scopes
   *
   * @returns {Promise<string>}
   * @memberof TeamsSSOProvider
   */
  public async showConsent(): Promise<string> {
    const teams = TeamsHelper.microsoftTeamsLib;

    return new Promise((resolve, reject) => {
      teams.getContext(context => {
        const msalAgent: UserAgentApplication = new UserAgentApplication({
          auth: { clientId: this._ssoConfig.clientId }
        });
        msalAgent
          .acquireTokenPopup({
            scopes: this._ssoConfig.scopes,
            prompt: 'consent',
            loginHint: context.loginHint
          })
          .then(result => {
            resolve(result.accessToken);
          })
          .catch(reason => {
            reject(reason);
          });
      });
    });
  }

  /**
   * Returns an access token that can be used for making calls to the Microsoft Graph
   *
   * @param {AuthenticationProviderOptions} options
   * @returns {Promise<string>}
   * @memberof TeamsSSOProvider
   */
  public async getAccessToken(options?: AuthenticationProviderOptions): Promise<string> {
    const scopes = options ? options.scopes || this._ssoConfig.scopes : this._ssoConfig.scopes;
    const url = new URL(this._ssoUrl, new URL(window.location.href));
    // Get token via SSO
    const clientToken = await this.getClientToken();

    // Exchange token from server
    const response: Response = await fetch(url.href, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        authorization: `Bearer ${clientToken}`
      },
      body: JSON.stringify({
        scopes: scopes,
        clientid: this._ssoConfig.clientId
      }),
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
  }

  /**
   * Get a token via the Teams SDK
   *
   * @returns {Promise<string>}
   * @memberof TeamsSSOProvider
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

  /**
   * Makes sure we can get an access token before considered logged in
   *
   * @returns {Promise<void>}
   * @memberof TeamsSSOProvider
   */
  private async internalLogin(): Promise<void> {
    let accessToken: string = await this.getAccessToken();

    // If we need to consent. Make sure we do this once during the log in process
    if (this._needsConsent && !accessToken) {
      // If consent is successful, we get an access token
      accessToken = await this.showConsent();
    }
    this.setState(accessToken ? ProviderState.SignedIn : ProviderState.SignedOut);
  }

  private unhandledFetchError(err: any) {
    console.error(`Unhandled fetch error: ${err}`);
  }
}
