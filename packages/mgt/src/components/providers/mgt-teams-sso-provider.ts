/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, property } from 'lit-element';
import { Providers } from '@microsoft/mgt-element';
import { MgtBaseProvider } from './baseProvider';
import { TeamsSSOConfig, TeamsSSOProvider } from '../../providers/TeamsSSOProvider';

/**
 * Authentication Library Provider for Microsoft Teams SSO accounts
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
    attribute: 'sso-url',
    type: String
  })
  public ssoUrl = '';

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
    if (this.clientId && this.ssoUrl && this.scopes) {
      const config: TeamsSSOConfig = {
        ssoUrl: this.ssoUrl,
        clientId: this.clientId,
        scopes: this.scopes.split(',')
      };

      this.provider = new TeamsSSOProvider(config);
      Providers.globalProvider = this.provider;
    }
  }
}
