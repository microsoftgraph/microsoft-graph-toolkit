/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IProvider, ProviderState, createFromProvider } from '@microsoft/mgt-element';
import { loadConfiguration, TeamsUserCredential } from '@microsoft/teamsfx';

/**
 * Interface to define the configuration when creating a TeamsFxConfigProvider
 *
 * @export
 * @interface TeamsFxConfig
 */
export interface TeamsFxConfig {
  /**
   * The app clientId
   *
   * @type {string}
   * @memberof TeamsFxConfig
   */
  clientId?: string;
  /**
   * Login page for Teams to redirect to.  Default value comes from REACT_APP_START_LOGIN_PAGE_URL environment variable.
   *
   * @type {string}
   * @memberof TeamsFxConfig
   */
  initiateLoginEndpoint?: string;
  /**
   * Endpoint of auth service provisioned by Teams Framework. Default value comes from REACT_APP_TEAMSFX_ENDPOINT environment variable.
   *
   * @type {string}
   * @memberof TeamsFxConfig
   */
  simpleAuthEndpoint?: string;
  /**
   * The scopes to use when authenticating the user
   *
   * @type {string[]}
   * @memberof TeamsFxConfig
   */
  scopes?: string[];
  /**
   * Credentials provided by TeamsFx
   *
   * @type {TeamsUserCredential}
   * @memberof TeamsFxConfig
   */
  credential?: TeamsUserCredential;
}

export declare interface AccessToken {
  /**
   * The access token returned by the authentication service.
   */
  token: string;
  /**
   * The access token's expiration timestamp in milliseconds, UNIX epoch time.
   */
  expiresOnTimestamp: number;
}

/**
 * TeamsFx Provider handler
 *
 * @export
 * @class TeamsFxProvider
 * @extends {IProvider}
 */
export class TeamsFxProvider extends IProvider {
  /**
   * Name used for analytics
   *
   * @readonly
   * @memberof TeamsFxProvider
   */
  public get name(): string {
    return 'MgtTeamsFxProvider';
  }

  /**
   * returns _credential
   *
   * @readonly
   * @memberof TeamsFxProvider
   */
  public get credential(): TeamsUserCredential {
    return this._credential;
  }

  /**
   * Privilege level for authentication
   *
   * @type {string[]}
   * @memberof TeamsFxProvider
   */
  private scopes: string[] = [];

  /**
   * Underlying TeamsFx credential
   *
   * @type {TeamsUserCredential}
   * @memberof TeamsFxProvider
   */
  private _credential: TeamsUserCredential;

  /**
   * Access token provided by TeamsFx
   *
   * @type {string}
   * @memberof TeamsFxProvider
   */
  private _accessToken: string = '';

  constructor(config: TeamsFxConfig) {
    super();
    this.scopes = config.scopes!;
    this._credential = config.credential!;

    loadConfiguration({
      authentication: {
        clientId: config.clientId ?? process.env.REACT_APP_CLIENT_ID,
        initiateLoginEndpoint: config.initiateLoginEndpoint ?? process.env.REACT_APP_START_LOGIN_PAGE_URL,
        simpleAuthEndpoint: config.simpleAuthEndpoint ?? process.env.REACT_APP_TEAMSFX_ENDPOINT
      }
    });

    if (!this._credential) {
      this._credential = new TeamsUserCredential();
    }

    this.graph = createFromProvider(this);
    this.internalLogin();
  }

  /**
   * Uses provider to receive access token via TeamsFx
   *
   * @returns {Promise<string>}
   * @memberof TeamsFxProvider
   */
  public async getAccessToken(): Promise<string> {
    let accessToken = '';
    accessToken = (await this.credential.getToken(this.scopes))!.token;
    return accessToken;
  }

  /**
   * Uses provider to receive access token via TeamsFx
   *
   * @returns {Promise<AccessToken>}
   * @memberof TeamsFxProvider
   */
  public async getFullAccessToken(): Promise<AccessToken> {
    let accessToken: AccessToken;
    accessToken = (await this.credential.getToken(this.scopes))!;
    return accessToken;
  }

  /**
   * Update scopes
   *
   * @param {string[]} scopes
   * @memberof TeamsFxProvider
   */
  public updateScopes(scopes: string[]) {
    this.scopes = scopes;
  }

  /**
   * Performs the internal login using TeamsFx
   *
   * @returns {Promise<void>}
   * @memberof TeamsFxProvider
   */
  private async internalLogin(): Promise<void> {
    let token: AccessToken | null = null;

    try {
      token = await this.getFullAccessToken();
    } catch (error) {
      await this.credential.login(this.scopes);
    }

    if (!!token && token.expiresOnTimestamp < +new Date()) {
      await this.credential.login(this.scopes);
    }

    this._accessToken = await this.getAccessToken();
    this.setState(this._accessToken ? ProviderState.SignedIn : ProviderState.SignedOut);
  }
}
