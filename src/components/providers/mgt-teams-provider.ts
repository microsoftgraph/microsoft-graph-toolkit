/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, LitElement, property } from 'lit-element';
import { Providers } from '../../Providers';
import { TeamsConfig, TeamsProvider } from '../../providers/TeamsProvider';
import { MgtBaseProvider } from './baseProvider';

/**
 * Authentication Library Provider for Microsoft Teams accounts
 *
 * @export
 * @class MgtTeamsProvider
 * @extends {MgtBaseProvider}
 */
@customElement('mgt-teams-provider')
export class MgtTeamsProvider extends MgtBaseProvider {
  /**
   * String alphanumerical value relation to a specific user
   *
   * @memberof MgtTeamsProvider
   */
  @property({
    attribute: 'client-id',
    type: String
  })
  public clientId = '';

  /**
   * The relative or absolute path of the html page that will handle the authentication
   *
   * @memberof MgtTeamsProvider
   */
  @property({
    attribute: 'auth-popup-url',
    type: String
  })
  public authPopupUrl = '';

  /**
   * Comma separated list of scopes.
   *
   * @memberof MgtTeamsProvider
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
   * @memberof MgtTeamsProvider
   */
  public get isAvailable() {
    return TeamsProvider.isAvailable;
  }
  /**
   * method called to initialize the provider. Each derived class should provide their own implementation
   *
   * @protected
   * @memberof MgtTeamsProvider
   */
  protected initializeProvider() {
    if (this.clientId && this.authPopupUrl) {
      const config: TeamsConfig = {
        authPopupUrl: this.authPopupUrl,
        clientId: this.clientId
      };

      if (this.scopes) {
        const scope = this.scopes.split(',');
        if (scope && scope.length > 0) {
          config.scopes = scope;
        }
      }

      this.provider = new TeamsProvider(config);
      Providers.globalProvider = this.provider;
    }
  }
}
