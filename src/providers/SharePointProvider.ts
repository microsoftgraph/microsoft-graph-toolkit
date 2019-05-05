/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IProvider, ProviderState } from './IProvider';
import { Graph } from '../Graph';

declare interface AadTokenProvider {
  getToken(x: string);
}

export declare interface WebPartContext {
  aadTokenProviderFactory: any;
}

export class SharePointProvider extends IProvider {
  private _idToken: string;

  private _provider: AadTokenProvider;

  get provider() {
    return this._provider;
  }

  get isLoggedIn(): boolean {
    return !!this._idToken;
  }

  private context: WebPartContext;

  scopes: string[];
  authority: string;

  constructor(context: WebPartContext) {
    super();
    this.context = context;

    context.aadTokenProviderFactory.getTokenProvider().then(
      (tokenProvider: AadTokenProvider): void => {
        this._provider = tokenProvider;
        this.graph = new Graph(this);
        this.internalLogin();
      }
    );
  }

  private async internalLogin(): Promise<void> {
    this._idToken = await this.getAccessToken();
    this.setState(this._idToken ? ProviderState.SignedIn : ProviderState.SignedOut);
  }

  async getAccessToken(): Promise<string> {
    let accessToken: string;
    try {
      accessToken = await this.provider.getToken('https://graph.microsoft.com');
    } catch (e) {
      console.log(e);
      throw e;
    }
    return accessToken;
  }

  updateScopes(scopes: string[]) {
    this.scopes = scopes;
  }
}
