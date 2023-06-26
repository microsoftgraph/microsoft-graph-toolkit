/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client';
import { createFromProvider } from '../Graph';
import { IProvider } from './IProvider';

/**
 * Facilitates create of new custom provider
 *
 * @export
 * @class SimpleProvider
 * @extends {IProvider}
 */
export class SimpleProvider extends IProvider {
  private readonly _getAccessTokenHandler: (scopes: string[]) => Promise<string>;
  private readonly _loginHandler: () => Promise<void>;
  private readonly _logoutHandler: () => Promise<void>;

  /**
   * Name used for analytics
   *
   * @readonly
   * @memberof IProvider
   */
  public get name() {
    return 'MgtSimpleProvider';
  }

  constructor(
    getAccessTokenHandler: (scopes: string[]) => Promise<string>,
    loginHandler?: () => Promise<void>,
    logoutHandler?: () => Promise<void>
  ) {
    super();

    this._getAccessTokenHandler = getAccessTokenHandler;
    this._loginHandler = loginHandler;
    this._logoutHandler = logoutHandler;

    this.graph = createFromProvider(this);
  }

  /**
   * Invokes the getAccessToken function
   *
   * @param {AuthenticationProviderOptions} [options]
   * @returns {Promise<string>}
   * @memberof SimpleProvider
   */
  public getAccessToken(options?: AuthenticationProviderOptions): Promise<string> {
    return this._getAccessTokenHandler(options?.scopes);
  }

  /**
   * Invokes login function
   *
   * @returns {Promise<void>}
   * @memberof SimpleProvider
   */
  public login(): Promise<void> {
    return this._loginHandler();
  }

  /**
   * Invokes logout function
   *
   * @returns {Promise<void>}
   * @memberof SimpleProvider
   */
  public logout(): Promise<void> {
    return this._logoutHandler();
  }
}
