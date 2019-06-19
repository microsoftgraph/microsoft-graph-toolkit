/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { LitElement, customElement, property } from 'lit-element';
import { Providers } from '../../Providers';
import { TeamsProvider, TeamsConfig } from '../../providers/TeamsProvider';

@customElement('mgt-teams-provider')
export class MgtTeamsProvider extends LitElement {
  @property({
    type: String,
    attribute: 'client-id'
  })
  clientId = '';

  @property({
    type: String,
    attribute: 'auth-popup-url'
  })
  authPopupUrl = '';

  /* Comma separated list of scopes. */
  @property({
    type: String,
    attribute: 'scopes'
  })
  scopes;

  firstUpdated(changedProperties) {
    if (TeamsProvider.isAvailable()) {
      this.validateAuthProps();
    }
  }

  private validateAuthProps() {
    if (this.clientId && this.authPopupUrl) {
      if (!Providers.globalProvider) {
        let config: TeamsConfig = {
          clientId: this.clientId,
          authPopupUrl: this.authPopupUrl
        };

        if (this.scopes) {
          let scope = this.scopes.split(',');
          if (scope && scope.length > 0) {
            config.scopes = scope;
          }
        }

        Providers.globalProvider = new TeamsProvider(config);
      }
    }
  }
}
