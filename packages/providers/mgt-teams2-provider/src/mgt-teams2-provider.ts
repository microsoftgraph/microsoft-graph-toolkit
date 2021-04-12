/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, property } from 'lit-element';
import { Configuration } from '@azure/msal-browser';
import { Providers, MgtBaseProvider } from '@microsoft/mgt-element';
import { Teams2Config, Teams2Provider } from './Teams2Provider';

/**
 * Authentication Library Provider for Microsoft Teams accounts
 *
 * @export
 * @class MgtTeams2Provider
 * @extends {MgtBaseProvider}
 */
@customElement('mgt-teams2-provider')
export class MgtTeams2Provider extends MgtBaseProvider {
  /**
   * String alphanumerical value relation to a specific user
   *
   * @memberof MgtTeams2Provider
   */
  @property({
    attribute: 'client-id',
    type: String
  })
  public clientId = '';

  /**
   * The relative or absolute path of the html page that will handle the authentication
   *
   * @memberof MgtTeams2Provider
   */
  @property({
    attribute: 'auth-popup-url',
    type: String
  })
  public authPopupUrl = '';

  /**
   * The relative or absolute path to the token exchange backend service
   *
   * @memberof MgtTeams2Provider
   */
  @property({
    attribute: 'sso-url',
    type: String
  })
  public ssoUrl = '';

  /**
   * The authority to use.
   *
   * @memberof MgtTeams2Provider
   */
  @property() public authority;

  /**
   * Comma separated list of scopes.
   *
   * @memberof MgtTeams2Provider
   */
  @property({
    attribute: 'scopes',
    type: String
  })
  public scopes;
  /**
   * Gets whether this provider can be used in this environment
   *
   * @readonly
   * @memberof MgtTeams2Provider
   */
  public get isAvailable() {
    return Teams2Provider.isAvailable;
  }
  /**
   * method called to initialize the provider. Each derived class should provide their own implementation
   *
   * @protected
   * @memberof MgtTeams2Provider
   */
  protected initializeProvider() {
    if (this.clientId && this.authPopupUrl) {
      const config: Teams2Config = {
        authPopupUrl: this.authPopupUrl,
        clientId: this.clientId
      };

      if (this.scopes) {
        const scope = this.scopes.split(',');
        if (scope && scope.length > 0) {
          config.scopes = scope;
        }
      }

      if (this.authority) {
        const msalConfig: Configuration = {
          auth: {
            authority: this.authority,
            clientId: this.clientId
          }
        };
        config.msalOptions = msalConfig;
      }

      if (this.ssoUrl) {
        config.ssoUrl = this.ssoUrl;
      }

      this.provider = new Teams2Provider(config);
      Providers.globalProvider = this.provider;
    }
  }
}
