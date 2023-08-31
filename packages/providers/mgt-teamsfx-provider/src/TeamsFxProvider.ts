/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  IProvider,
  ProviderState,
  createFromProvider,
  GraphEndpoint,
  MICROSOFT_GRAPH_DEFAULT_ENDPOINT
} from '@microsoft/mgt-element';
import { TokenCredential } from '@azure/core-auth';

/**
 * Interface represents TeamsUserCredential in TeamsFx library
 */
export interface TeamsFxUserCredential extends TokenCredential {
  login(scopes: string | string[], resources?: string[]): Promise<void>;
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
   * Privilege level for authentication
   *
   * Can use string array or space-separated string, such as ["User.Read", "Application.Read.All"] or "User.Read Application.Read.All"
   *
   * @type {string | string[]}
   * @memberof TeamsFxProvider
   */
  // eslint-disable-next-line @typescript-eslint/naming-convention
  private readonly scopes: string | string[] = [];

  /**
   * TeamsFxUserCredential instance
   *
   * @type {TeamsFx}
   * @memberof TeamsFxProvider
   */
  private readonly _credential: TeamsFxUserCredential;

  /**
   * Access token provided by TeamsFx
   *
   * @type {string}
   * @memberof TeamsFxProvider
   */
  private _accessToken = '';

  /**
   * Constructor of TeamsFxProvider.
   *
   * @example
   * ```typescript
   * import {Providers} from '@microsoft/mgt-element';
   * import {TeamsFxProvider} from '@microsoft/mgt-teamsfx-provider';
   * import {TeamsUserCredential, TeamsUserCredentialAuthConfig} from "@microsoft/teamsfx";
   *
   * const authConfig: TeamsUserCredentialAuthConfig = {
   *     clientId: process.env.REACT_APP_CLIENT_ID,
   *     initiateLoginEndpoint: process.env.REACT_APP_START_LOGIN_PAGE_URL,
   * };
   * const scope = ["User.Read"];
   *
   * const credential = new TeamsUserCredential(authConfig);
   * const provider = new TeamsFxProvider(credential, scope);
   * Providers.globalProvider = provider;
   * ```
   *
   * @param {TeamsFxUserCredential} credential - TeamsUserCredential instance in TeamsFx library.
   * @param {string | string[]} scopes - The list of scopes for which the token will have access.
   * @param {GraphEndpoint} baseURL - Graph endpoint.
   *
   */
  constructor(
    credential: TeamsFxUserCredential,
    scopes: string | string[],
    baseURL: GraphEndpoint = MICROSOFT_GRAPH_DEFAULT_ENDPOINT
  ) {
    super();

    if (!this._credential) {
      this._credential = credential;
    }

    this.validateScopesType(scopes);

    const scopesArr = this.getScopesArray(scopes);

    if (!scopesArr || scopesArr.length === 0) {
      this.scopes = ['.default'];
    } else {
      this.scopes = scopesArr;
    }
    this.approvedScopes = this.scopes;

    this.baseURL = baseURL;
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
      const accessToken = await this._credential.getToken(this.scopes);
      this._accessToken = accessToken ? accessToken.token : '';
      if (!this._accessToken) {
        throw new Error('Access token is null');
      }
    } catch (error: unknown) {
      const err = error as object;
      // eslint-disable-next-line no-console
      console.error(`🦒: Cannot get access token due to error: ${err.toString()}`);
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
      await this._credential.login(this.scopes);
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
