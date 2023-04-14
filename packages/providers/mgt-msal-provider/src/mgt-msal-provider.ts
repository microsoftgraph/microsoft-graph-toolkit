/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
import { property } from 'lit/decorators.js';
import { Providers, LoginType, MgtBaseProvider, customElement } from '@microsoft/mgt-element';
import { MsalConfig, MsalProvider } from './MsalProvider';
/**
 * Authentication Library Provider for Microsoft personal accounts
 *
 * @export
 * @class MgtMsalProvider
 * @extends {MgtBaseProvider}
 */
@customElement('msal-provider')
// @customElement('mgt-msal-provider')
export class MgtMsalProvider extends MgtBaseProvider {
  /**
   * String alphanumerical value relation to a specific user
   *
   * @memberof MgtMsalProvider
   */
  @property({
    attribute: 'client-id',
    type: String
  })
  public clientId = '';

  /**
   * The login type that should be used: popup or redirect
   *
   * @memberof MgtMsalProvider
   */
  @property({
    attribute: 'login-type',
    type: String
  })
  public loginType: string;

  /**
   * The authority to use.
   *
   * @memberof MgtMsalProvider
   */
  @property() public authority: string;

  /**
   * Comma separated list of scopes
   *
   * @memberof MgtMsalProvider
   */
  @property({
    attribute: 'scopes',
    type: String
  })
  public scopes: string;

  /**
   * The redirect uri to use
   *
   * @memberof MgtMsalProvider
   */
  @property({
    attribute: 'redirect-uri',
    type: String
  })
  public redirectUri: string;

  /**
   * The domain hint to use during login
   *
   * @memberof MgtMsalProvider
   */
  @property({
    attribute: 'domain-hint',
    type: String
  })
  public domainHint: string;

  /**
   * The prompt type to use during login
   *
   * @memberof MgtMsalProvider
   */
  @property({
    attribute: 'prompt',
    type: String
  })
  public prompt: string;

  /**
   * Gets whether this provider can be used in this environment
   *
   * @readonly
   * @memberof MgtMsalProvider
   */
  public get isAvailable() {
    return true;
  }

  /**
   * method called to initialize the provider. Each derived class should provide their own implementation.
   *
   * @protected
   * @memberof MgtMsalProvider
   */
  protected initializeProvider() {
    if (this.clientId) {
      const config: MsalConfig = {
        clientId: this.clientId,
        domainHint: this.domainHint,
        prompt: this.prompt
      };

      if (this.loginType && this.loginType.length > 1) {
        let loginType: string = this.loginType.toLowerCase();
        loginType = loginType[0].toUpperCase() + loginType.slice(1);
        const loginTypeEnum = LoginType[loginType] as LoginType;
        config.loginType = loginTypeEnum;
      }

      if (this.authority) {
        config.authority = this.authority;
      }

      if (this.scopes) {
        const scope = this.scopes.split(',');
        if (scope && scope.length > 0) {
          config.scopes = scope;
        }
      }

      if (this.redirectUri) {
        config.redirectUri = this.redirectUri;
      }

      if (this.baseUrl) {
        config.baseURL = this.baseUrl;
      }

      this.provider = new MsalProvider(config);
      Providers.globalProvider = this.provider;
    }
  }
}
