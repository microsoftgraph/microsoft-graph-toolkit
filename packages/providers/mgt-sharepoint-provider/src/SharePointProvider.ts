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

/**
 * AadTokenProvider
 *
 * @interface AadTokenProvider
 */
declare interface AadTokenProvider {
  /**
   * get token with x
   *
   * @param {string} x
   * @memberof AadTokenProvider
   */
  getToken(x: string): Promise<string>;
}

/**
 * contains the contextual services available to a web part
 *
 * @export
 * @interface WebPartContext
 */
declare interface WebPartContext {
  aadTokenProviderFactory: { getTokenProvider(): Promise<AadTokenProvider> };
}

/**
 * SharePoint Provider handler
 *
 * @export
 * @class SharePointProvider
 * @extends {IProvider}
 */
export class SharePointProvider extends IProvider {
  /**
   * returns _provider
   *
   * @readonly
   * @memberof SharePointProvider
   */
  get provider(): AadTokenProvider {
    return this._provider;
  }

  /**
   * returns _idToken
   *
   * @readonly
   * @type {boolean}
   * @memberof SharePointProvider
   */
  get isLoggedIn(): boolean {
    return !!this._idToken;
  }

  /**
   * Name used for analytics
   *
   * @readonly
   * @memberof IProvider
   */
  public get name() {
    return 'MgtSharePointProvider';
  }

  /**
   * privilege level for authentication
   *
   * @type {string[]}
   * @memberof SharePointProvider
   */
  public scopes: string[];

  /**
   * authority
   *
   * @type {string}
   * @memberof SharePointProvider
   */
  public authority: string;
  private _idToken: string;

  private _provider: AadTokenProvider;

  constructor(context: WebPartContext, baseUrl: GraphEndpoint = MICROSOFT_GRAPH_DEFAULT_ENDPOINT) {
    super();

    void context.aadTokenProviderFactory.getTokenProvider().then((tokenProvider: AadTokenProvider): void => {
      this._provider = tokenProvider;
      this.baseURL = baseUrl;
      this.graph = createFromProvider(this);
      void this.internalLogin();
    });
  }

  /**
   * uses provider to receive access token via SharePoint Provider
   *
   * @returns {Promise<string>}
   * @memberof SharePointProvider
   */
  public async getAccessToken(): Promise<string> {
    const baseUrl = this.baseURL ? this.baseURL : MICROSOFT_GRAPH_DEFAULT_ENDPOINT;
    return await this.provider.getToken(baseUrl);
  }
  /**
   * update scopes
   *
   * @param {string[]} scopes
   * @memberof SharePointProvider
   */
  public updateScopes(scopes: string[]) {
    this.scopes = scopes;
  }

  private async internalLogin(): Promise<void> {
    this._idToken = await this.getAccessToken();
    this.setState(this._idToken ? ProviderState.SignedIn : ProviderState.SignedOut);
  }
}
