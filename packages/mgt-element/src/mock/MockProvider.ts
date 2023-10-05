/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IProvider, IProviderAccount, ProviderState } from '../providers/IProvider';
import { MockGraph } from './MockGraph';

/**
 * Mock Provider access token for Microsoft Graph APIs
 *
 * @export
 * @class MockProvider
 * @extends {IProvider}
 */
export class MockProvider extends IProvider {
  public provider: IProvider;

  private readonly _mockGraphPromise: Promise<MockGraph>;

  /**
   * new instance of mock graph provider
   *
   * @memberof MockProvider
   */
  public graph: MockGraph;
  constructor(signedIn = false, signedInAccounts: IProviderAccount[] = []) {
    super();
    this._mockGraphPromise = MockGraph.create(this);
    const enableMultipleLogin = Boolean(signedInAccounts.length);
    this.isMultipleAccountSupported = enableMultipleLogin;
    this.isMultipleAccountDisabled = !enableMultipleLogin;
    this._accounts = signedInAccounts;

    void this.initializeMockGraph(signedIn);
  }

  /**
   * Indicates if the MockProvider is configured to support multi account mode
   * This is only true if the Mock provider has been configured with signedInAccounts in the constructor
   *
   * @readonly
   * @type {boolean}
   * @memberof MockProvider
   */
  public get isMultiAccountSupportedAndEnabled(): boolean {
    return !this.isMultipleAccountDisabled && this.isMultipleAccountSupported;
  }

  private readonly _accounts: IProviderAccount[] = [];
  /**
   * Returns the array of accounts the MockProviders has been configured with
   *
   * @return {*}  {IProviderAccount[]}
   * @memberof MockProvider
   */
  public getAllAccounts?(): IProviderAccount[] {
    return this._accounts;
  }

  /**
   * Returns the first account in the set of accounts the MockProvider has been configured with
   *
   * @return {*}  {IProviderAccount}
   * @memberof MockProvider
   */
  public getActiveAccount?(): IProviderAccount {
    if (this._accounts.length) {
      return this._accounts[0];
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
    await this._mockGraphPromise;
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
    await this._mockGraphPromise;
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

  /**
   * Name used for analytics
   *
   * @readonly
   * @memberof IProvider
   */
  public get name() {
    return 'MgtMockProvider';
  }

  private async initializeMockGraph(signedIn = false) {
    this.graph = await this._mockGraphPromise;

    if (signedIn) {
      this.setState(ProviderState.SignedIn);
    } else {
      this.setState(ProviderState.SignedOut);
    }
  }
}
