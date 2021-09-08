/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, property } from 'lit-element';
import { Configuration } from '@azure/msal-browser';
import { Providers, MgtBaseProvider } from '@microsoft/mgt-element';
import { HttpMethod, TeamsMsal2Config, TeamsMsal2Provider } from './TeamsMsal2Provider';

/**
 * Authentication Library Provider for Microsoft Teams accounts
 *
 * @export
 * @class MgtTeamsMsal2Provider
 * @extends {MgtBaseProvider}
 */
@customElement('mgt-teams-msal2-provider')
export class MgtTeamsMsal2Provider extends MgtBaseProvider {
  /**
   * String alphanumerical value relation to a specific user
   *
   * @memberof MgtTeamsMsal2Provider
   */
  @property({
    attribute: 'client-id',
    type: String
  })
  public clientId = '';

  /**
   * The relative or absolute path of the html page that will handle the authentication
   *
   * @memberof MgtTeamsMsal2Provider
   */
  @property({
    attribute: 'auth-popup-url',
    type: String
  })
  public authPopupUrl = '';

  /**
   * The relative or absolute path to the token exchange backend service
   *
   * @memberof MgtTeamsMsal2Provider
   */
  @property({
    attribute: 'sso-url',
    type: String
  })
  public ssoUrl = '';

  /**
   * The authority to use.
   *
   * @memberof MgtTeamsMsal2Provider
   */
  @property() public authority;

  /**
   * Comma separated list of scopes.
   *
   * @memberof MgtTeamsMsal2Provider
   */
  @property({
    attribute: 'scopes',
    type: String
  })
  public scopes;
  /**
   * Disables auto display of popup when consent is required
   *
   * @memberof MgtTeamsMsal2Provider
   */
  @property({
    attribute: 'auto-consent-disabled',
    type: Boolean
  })
  public isAutoConsentDisabled;
  /**
   * Disables auto display of popup when consent is required
   *
   * @memberof MgtTeamsMsal2Provider
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
   * @memberof MgtTeamsMsal2Provider
   */
  public get isAvailable() {
    return TeamsMsal2Provider.isAvailable;
  }
  /**
   * method called to initialize the provider. Each derived class should provide their own implementation
   *
   * @protected
   * @memberof MgtTeamsMsal2Provider
   */
  protected initializeProvider() {
    if (this.clientId && this.authPopupUrl) {
      const config: TeamsMsal2Config = {
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

      this.provider = new TeamsMsal2Provider(config);
      Providers.globalProvider = this.provider;
    }
  }
}
