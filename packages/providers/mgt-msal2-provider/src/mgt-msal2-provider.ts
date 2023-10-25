/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { property } from 'lit/decorators.js';
import { Providers, LoginType, MgtBaseProvider, registerComponent } from '@microsoft/mgt-element';
import { Msal2Config, Msal2Provider, PromptType } from './Msal2Provider';

export const registerMgtMsal2Provider = () => {
  registerComponent('msal2-provider', MgtMsal2Provider);
};

/**
 * Authentication Library Provider for Microsoft personal accounts
 *
 * @export
 * @class MgtMsalProvider
 * @extends {MgtBaseProvider}
 */
class MgtMsal2Provider extends MgtBaseProvider {
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
   * login hint string
   *
   * @memberof MgtMsal2Provider
   */
  @property({
    attribute: 'login-hint',
    type: String
  })
  public loginHint: string;

  /**
   * domain hint string
   *
   * @memberof MgtMsal2Provider
   */
  @property({
    attribute: 'domain-hint',
    type: String
  })
  public domainHint: string;

  /**
   * The login type that should be used: popup or redirect
   *
   * @memberof MgtMsal2Provider
   */
  @property({
    attribute: 'login-type',
    type: String
  })
  public loginType: string;

  /**
   * The authority to use.
   *
   * @memberof MgtMsal2Provider
   */
  @property()
  public authority: string;

  /**
   * Comma separated list of scopes
   *
   * @memberof MgtMsal2Provider
   */
  @property({
    attribute: 'scopes',
    type: String
  })
  public scopes: string;

  /**
   * The redirect uri to use
   *
   * @memberof MgtMsal2Provider
   */
  @property({
    attribute: 'redirect-uri',
    type: String
  })
  public redirectUri: string;

  /**
   * Type of prompt for login
   *
   * @memberof MgtMsal2Provider
   */
  @property({
    attribute: 'prompt',
    type: String
  })
  public prompt: string;

  /**
   * Disables incremental consent
   *
   * @memberof MgtMsal2Provider
   */
  @property({
    attribute: 'incremental-consent-disabled',
    type: Boolean
  })
  public isIncrementalConsentDisabled: boolean;

  /**
   * Disables multiple account capability
   *
   * @memberof MgtMsal2Provider
   */
  @property({
    attribute: 'multi-account-disabled',
    type: Boolean
  })
  public isMultiAccountDisabled;

  /**
   * Gets whether this provider can be used in this environment
   *
   * @readonly
   * @memberof MgtMsal2Provider
   */
  public get isAvailable() {
    return true;
  }

  /**
   * method called to initialize the provider. Each derived class should provide their own implementation.
   *
   * @protected
   * @memberof MgtMsal2Provider
   */
  protected initializeProvider() {
    if (this.clientId) {
      const config: Msal2Config = {
        clientId: this.clientId
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

      if (this.loginHint) {
        config.loginHint = this.loginHint;
      }

      if (this.domainHint) {
        config.domainHint = this.domainHint;
      }

      if (this.prompt) {
        const prompt: string = this.prompt.toUpperCase();
        const promptEnum = PromptType[prompt] as PromptType;
        config.prompt = promptEnum;
      }

      if (this.isIncrementalConsentDisabled) {
        config.isIncrementalConsentDisabled = true;
      }

      if (this.isMultiAccountDisabled) {
        config.isMultiAccountEnabled = false;
      }

      if (this.baseUrl) {
        config.baseURL = this.baseUrl;
      }

      if (this.customHosts) {
        config.customHosts = this.customHosts;
      }

      this.provider = new Msal2Provider(config);
      Providers.globalProvider = this.provider;
    }
  }
}
