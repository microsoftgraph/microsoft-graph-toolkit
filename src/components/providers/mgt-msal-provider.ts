/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { LitElement, customElement, property } from 'lit-element';
import { MsalConfig, MsalProvider } from '../../providers/MsalProvider';
import { LoginType } from '../../providers/IProvider';
import { Providers } from '../../Providers';

@customElement('mgt-msal-provider')
export class MgtMsalProvider extends LitElement {
  private _isInitialized: boolean = false;

  @property({
    type: String,
    attribute: 'client-id'
  })
  clientId = '';

  @property({
    type: String,
    attribute: 'login-type'
  })
  loginType;

  @property() authority;

  /* Comma separated list of scopes. */
  @property({
    type: String,
    attribute: 'scopes'
  })
  scopes;

  firstUpdated(changedProperties) {
    this.validateAuthProps();
  }

  private validateAuthProps() {
    if (this._isInitialized) {
      return;
    }

    if (this.clientId) {
      this._isInitialized = true;

      let config: MsalConfig = {
        clientId: this.clientId
      };

      if (this.loginType && this.loginType.length > 1) {
        let loginType: string = this.loginType.toLowerCase();
        loginType = loginType[0].toUpperCase() + loginType.slice(1);
        let loginTypeEnum = LoginType[loginType];
        config.loginType = loginTypeEnum;
      }

      if (this.authority) {
        config.authority = this.authority;
      }

      if (this.scopes) {
        let scope = this.scopes.split(',');
        if (scope && scope.length > 0) {
          config.scopes = scope;
        }
      }

      Providers.globalProvider = new MsalProvider(config);
    }
  }
}
