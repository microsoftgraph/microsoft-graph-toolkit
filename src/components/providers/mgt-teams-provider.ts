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

@customElement('mgt-teams-provider')
export class MgtTeamsProvider extends MgtBaseProvider {
  @property({
    type: String,
    attribute: 'client-id'
  })
  public clientId = '';

  @property({
    type: String,
    attribute: 'auth-popup-url'
  })
  public authPopupUrl = '';

  /* Comma separated list of scopes. */
  @property({
    type: String,
    attribute: 'scopes'
  })
  public scopes;

  public get isAvailable() {
    return TeamsProvider.isAvailable;
  }

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
