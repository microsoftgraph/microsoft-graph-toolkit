/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IProvider, ProviderState } from '@microsoft/mgt-element';
import { MockGraph } from './MockGraph';

/**
 * Mock Provider access token for Microsoft Graph APIs
 *
 * @export
 * @class MockProvider
 * @extends {IProvider}
 */
export class MockProvider extends IProvider {
  // tslint:disable-next-line: completed-docs
  public provider: any;

  /**
   * new instance of mock graph provider
   *
   * @memberof MockProvider
   */
  public graph = new MockGraph(this);
  constructor(signedIn: boolean = false) {
    super();
    if (signedIn) {
      this.setState(ProviderState.SignedIn);
    } else {
      this.setState(ProviderState.SignedOut);
    }
  }

  /**
   * sets Provider state to SignedIn
   *
   * @returns {Promise<void>}
   * @memberof MockProvider
   */
  public async login(): Promise<void> {
    this.setState(ProviderState.Loading);
    await new Promise(resolve => setTimeout(resolve, 3000));
    this.setState(ProviderState.SignedIn);
  }
  /**
   * sets Provider state to signed out
   *
   * @returns {Promise<void>}
   * @memberof MockProvider
   */
  public async logout(): Promise<void> {
    this.setState(ProviderState.Loading);
    await new Promise(resolve => setTimeout(resolve, 3000));
    this.setState(ProviderState.SignedOut);
  }
  /**
   * Promise returning token from graph.microsoft.com
   *
   * @returns {Promise<string>}
   * @memberof MockProvider
   */
  public getAccessToken(): Promise<string> {
    return Promise.resolve('{token:https://graph.microsoft.com/}');
  }
}
