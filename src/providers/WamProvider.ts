/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';
import { Graph } from '../Graph';
import { IProvider, ProviderState } from './IProvider';

/**
 * Windows auth
 *
 * @interface Window
 */
declare interface Window {
  // tslint:disable-next-line: completed-docs
  Windows: any;
}

// tslint:disable-next-line: completed-docs
declare var window: Window;

/**
 * Wam provider authentication
 *
 * @export
 * @class WamProvider
 * @extends {IProvider}
 */
export class WamProvider extends IProvider {
  private graphResource = 'https://graph.microsoft.com';
  private clientId: string;
  private authority: string;

  private accessToken: string;

  /**
   * returns if Windows is the OS based on window
   *
   * @readonly
   * @static
   * @type {boolean}
   * @memberof WamProvider
   */
  public static get isAvailable(): boolean {
    return !!window.Windows;
  }

  /**
   * returns accessToken
   *
   * @readonly
   * @type {boolean}
   * @memberof WamProvider
   */
  get isLoggedIn(): boolean {
    return !!this.accessToken;
  }

  constructor(clientId: string, authority?: string) {
    super();
    this.clientId = clientId;
    this.authority = authority || 'https://login.microsoftonline.com/common';

    this.graph = new Graph(this);

    this.printRedirectUriToConsole();
  }

  /**
   * login promise for the WamProvider
   *
   * @returns {Promise<void>}
   * @memberof WamProvider
   */
  public async login(): Promise<void> {
    if (WamProvider.isAvailable) {
      const webCore = window.Windows.Security.Authentication.Web.Core;
      const wap = await webCore.WebAuthenticationCoreManager.findAccountProviderAsync(
        'https://login.microsoft.com',
        this.authority
      );
      if (!wap) {
        // tslint:disable-next-line: no-console
        console.log('no account provider');
        return;
      }

      const wtr = new webCore.WebTokenRequest(wap, '', this.clientId);
      wtr.properties.insert('resource', this.graphResource);

      const wtrr = await webCore.WebAuthenticationCoreManager.requestTokenAsync(wtr);
      switch (wtrr.responseStatus) {
        case webCore.WebTokenRequestStatus.success:
          const account = wtrr.responseData[0].webAccount;
          this.accessToken = wtrr.responseData[0].token;
          this.setState(this.accessToken ? ProviderState.SignedIn : ProviderState.SignedOut);
          break;
        case webCore.WebTokenRequestStatus.userCancel:
        case webCore.WebTokenRequestStatus.accountSwitch:
        case webCore.WebTokenRequestStatus.userInteractionRequired:
        case webCore.WebTokenRequestStatus.accountProviderNotAvailable:
        case webCore.WebTokenRequestStatus.providerError:
          // tslint:disable-next-line: no-console
          console.log(
            `status ${wtrr.responseStatus}: error code ${wtrr.responseError} | error message ${
              wtrr.responseError.errorMessage
            }`
          );
          break;
      }
    }
  }

  /**
   * console logs the redirect url for auth
   *
   * @memberof WamProvider
   */
  public printRedirectUriToConsole() {
    if (WamProvider.isAvailable) {
      const web = window.Windows.Security.Authentication.Web;
      const redirectUri = `ms-appx-web://Microsoft.AAD.BrokerPlugIn/${(web.WebAuthenticationBroker.getCurrentApplicationCallbackUri()
        .host as string).toUpperCase()}`;
      // tslint:disable-next-line: no-console
      console.log('Use the following redirect URI in your AAD application:');
      // tslint:disable-next-line: no-console
      console.log(redirectUri);
    } else {
      // tslint:disable-next-line: no-console
      console.log('WAM not supported on this platform');
    }
  }

  /**
   * returns access token Promise
   *
   * @param {AuthenticationProviderOptions} options
   * @returns {Promise<string>}
   * @memberof WamProvider
   */
  public getAccessToken(options: AuthenticationProviderOptions): Promise<string> {
    if (this.isLoggedIn) {
      return Promise.resolve(this.accessToken);
    } else {
      throw new Error('Not logged in');
    }
  }
}
