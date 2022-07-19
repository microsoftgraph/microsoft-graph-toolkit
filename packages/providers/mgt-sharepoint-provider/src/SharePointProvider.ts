/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IProvider, ProviderState, createFromProvider } from '@microsoft/mgt-element';

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
  getToken(x: string);
}

/**
 * contains the contextual services available to a web part
 *
 * @export
 * @interface WebPartContext
 */
declare interface WebPartContext {
  // tslint:disable-next-line: completed-docs
  aadTokenProviderFactory: any;
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
  get provider() {
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

  constructor(context: WebPartContext) {
    super();

    context.aadTokenProviderFactory.getTokenProvider().then((tokenProvider: AadTokenProvider): void => {
      this._provider = tokenProvider;
      this.graph = createFromProvider(this);
      this.internalLogin();
    });
  }

  /**
   * uses provider to receive access token via SharePoint Provider
   *
   * @returns {Promise<string>}
   * @memberof SharePointProvider
   */
  public async getAccessToken(): Promise<string> {
    let accessToken: string;
    try {
      accessToken = await this.provider.getToken('https://graph.microsoft.com');
    } catch (e) {
      throw e;
    }
    return accessToken;
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
