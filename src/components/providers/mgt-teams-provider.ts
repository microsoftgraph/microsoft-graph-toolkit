/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { LitElement, customElement, property } from 'lit-element';
import { Providers } from '../../Providers';
import { TeamsProvider } from '../../providers/TeamsProvider';

@customElement('mgt-teams-provider')
export class MgtTeamsProvider extends LitElement {
  private _provider: TeamsProvider;

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

  async firstUpdated(changedProperties) {
    this.validateAuthProps();
    if (await TeamsProvider.isAvailable()) {
      Providers.globalProvider = this._provider;
    }
  }

  private validateAuthProps() {
    if (this.clientId && this.authPopupUrl) {
      if (!this._provider) {
        this._provider = new TeamsProvider({
          clientId: this.clientId,
          authPopupUrl: this.authPopupUrl
        });
      }
    }
  }
}
