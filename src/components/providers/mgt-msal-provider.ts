/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, LitElement, property } from 'lit-element';
import { Providers } from '../../Providers';
import { LoginType } from '../../providers/IProvider';
import { MsalConfig, MsalProvider } from '../../providers/MsalProvider';
import { MgtBaseProvider } from './baseProvider';

@customElement('mgt-msal-provider')
export class MgtMsalProvider extends MgtBaseProvider {
  @property({
    type: String,
    attribute: 'client-id'
  })
  public clientId = '';

  @property({
    type: String,
    attribute: 'login-type'
  })
  public loginType;

  @property() public authority;

  /* Comma separated list of scopes. */
  @property({
    type: String,
    attribute: 'scopes'
  })
  public scopes;

  public get isAvailable() {
    return true;
  }

  protected initializeProvider() {
    if (this.clientId) {
      const config: MsalConfig = {
        clientId: this.clientId
      };

      if (this.loginType && this.loginType.length > 1) {
        let loginType: string = this.loginType.toLowerCase();
        loginType = loginType[0].toUpperCase() + loginType.slice(1);
        const loginTypeEnum = LoginType[loginType];
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

      this.provider = new MsalProvider(config);
      Providers.globalProvider = this.provider;
    }
  }
}
