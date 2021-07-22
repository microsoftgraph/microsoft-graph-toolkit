/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, property } from 'lit-element';
import { Configuration } from '@azure/msal-browser';
import { Providers, MgtBaseProvider } from '@microsoft/mgt-element';
import { HttpMethod, TeamsSSOConfig, TeamsSSOProvider } from './TeamsSSOProvider';

/**
 * Authentication Library Provider for Microsoft Teams accounts
 *
 * @export
 * @class MgtTeamsSSOProvider
 * @extends {MgtBaseProvider}
 */
@customElement('mgt-teams-sso-provider')
export class MgtTeamsSSOProvider extends MgtBaseProvider {
  /**
   * String alphanumerical value relation to a specific user
   *
   * @memberof MgtTeamsSSOProvider
   */
  @property({
    attribute: 'client-id',
    type: String
  })
  public clientId = '';

  /**
   * The relative or absolute path of the html page that will handle the authentication
   *
   * @memberof MgtTeamsSSOProvider
   */
  @property({
    attribute: 'auth-popup-url',
    type: String
  })
  public authPopupUrl = '';

  /**
   * The relative or absolute path to the token exchange backend service
   *
   * @memberof MgtTeamsSSOProvider
   */
  @property({
    attribute: 'sso-url',
    type: String
  })
  public ssoUrl = '';

  /**
   * The authority to use.
   *
   * @memberof MgtTeamsSSOProvider
   */
  @property() public authority;

  /**
   * Comma separated list of scopes.
   *
   * @memberof MgtTeamsSSOProvider
   */
  @property({
    attribute: 'scopes',
    type: String
  })
  public scopes;
  /**
   * Disables auto display of popup when consent is required
   *
   * @memberof MgtTeamsSSOProvider
   */
  @property({
    attribute: 'auto-consent-disabled',
    type: Boolean
  })
  public isAutoConsentDisabled;
  /**
   * Disables auto display of popup when consent is required
   *
   * @memberof MgtTeamsSSOProvider
   */
  @property({
    attribute: 'http-method',
    type: String
  })
  public httpMethod;
  /**
   * Gets whether this provider can be used in this environment
   *
   * @readonly
   * @memberof MgtTeamsSSOProvider
   */
  public get isAvailable() {
    return TeamsSSOProvider.isAvailable;
  }
  /**
   * method called to initialize the provider. Each derived class should provide their own implementation
   *
   * @protected
   * @memberof MgtTeamsSSOProvider
   */
  protected initializeProvider() {
    if (this.clientId && this.authPopupUrl) {
      const config: TeamsSSOConfig = {
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

      if (this.isAutoConsentDisabled) {
        config.autoConsent = false;
      }

      if (this.httpMethod) {
        let httpMethod: string = this.httpMethod.toUpperCase();
        const httpMethodEnum = HttpMethod[httpMethod];
        config.httpMethod = httpMethodEnum;
      }

      this.provider = new TeamsSSOProvider(config);
      Providers.globalProvider = this.provider;
    }
  }
}
