/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IProvider, ProviderState, createFromProvider } from '@microsoft/mgt-element';
import { TeamsUserCredential } from '@microsoft/teamsfx';

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

  constructor(credential: TeamsUserCredential, scopes: string[]) {
    super();

    if (!this._credential) {
      this._credential = credential;
    }

    if (!scopes || scopes.length === 0) {
      this.scopes = ['.default'];
    } else {
      this.scopes = scopes;
    }

    this.graph = createFromProvider(this);
  }

  /**
   * Uses provider to receive access token via TeamsFx
   *
   * @returns {Promise<string>}
   * @memberof TeamsFxProvider
   */
  public async getAccessToken(): Promise<string> {
    try {
      const accessToken = await this.credential.getToken(this.scopes);
      this._accessToken = accessToken ? accessToken.token : '';
      if (!this._accessToken) {
        throw new Error('Access token is null');
      }
    } catch (error) {
      this.setState(ProviderState.SignedOut);
      this._accessToken = '';
    }
    return this._accessToken;
  }

  /**
   * Performs the login using TeamsFx
   *
   * @returns {Promise<void>}
   * @memberof TeamsFxProvider
   */
  public async login(): Promise<void> {
    const token: string = await this.getAccessToken();

    if (!token) {
      await this.credential.login(this.scopes);
    }

    this._accessToken = token ?? (await this.getAccessToken());
    this.setState(this._accessToken ? ProviderState.SignedIn : ProviderState.SignedOut);
  }
}
