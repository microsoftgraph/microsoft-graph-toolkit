/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IProvider, ProviderState, createFromProvider } from '@microsoft/mgt-element';
import { TeamsFx } from '@microsoft/teamsfx';

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
   * returns teamsfx instance
   *
   * @readonly
   * @memberof TeamsFxProvider
   */
  public get teamsfx(): TeamsFx {
    return this._teamsfx;
  }

  /**
   * Privilege level for authentication
   *
   * Can use string array or space-separated string, such as ["User.Read", "Application.Read.All"] or "User.Read Application.Read.All"
   *
   * @type {string | string[]}
   * @memberof TeamsFxProvider
   */
  private scopes: string | string[] = [];

  /**
   * TeamsFx instance
   *
   * @type {TeamsFx}
   * @memberof TeamsFxProvider
   */
  private readonly _teamsfx: TeamsFx;

  /**
   * Access token provided by TeamsFx
   *
   * @type {string}
   * @memberof TeamsFxProvider
   */
  private _accessToken: string = '';

  constructor(teamsfx: TeamsFx, scopes: string | string[]) {
    super();

    if (!this._teamsfx) {
      this._teamsfx = teamsfx;
    }

    this.validateScopesType(scopes);

    const scopesArr = this.getScopesArray(scopes);

    if (!scopesArr || scopesArr.length === 0) {
      this.scopes = ['.default'];
    } else {
      this.scopes = scopesArr;
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
      const accessToken = await this.teamsfx.getCredential().getToken(this.scopes);
      this._accessToken = accessToken ? accessToken.token : '';
      if (!this._accessToken) {
        throw new Error('Access token is null');
      }
    } catch (error) {
      console.error('Cannot get access token due to error: ' + error.toString());
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
      await this.teamsfx.login(this.scopes);
    }

    this._accessToken = token ?? (await this.getAccessToken());
    this.setState(this._accessToken ? ProviderState.SignedIn : ProviderState.SignedOut);
  }

  private validateScopesType(value: any): void {
    // string
    if (typeof value === 'string' || value instanceof String) {
      return;
    }

    // empty array
    if (Array.isArray(value) && value.length === 0) {
      return;
    }

    // string array
    if (Array.isArray(value) && value.length > 0 && value.every(item => typeof item === 'string')) {
      return;
    }

    throw new Error('The type of scopes is not valid, it must be string or string array');
  }

  private getScopesArray(scopes: string | string[]): string[] {
    const scopesArray: string[] = typeof scopes === 'string' ? scopes.split(' ') : scopes;
    return scopesArray.filter(x => x !== null && x !== '');
  }
}
