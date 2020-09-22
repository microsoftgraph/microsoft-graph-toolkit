// update these values with ones your AAD Client Id

import { MsalAngularConfiguration } from '@azure/msal-angular';
import { Configuration } from 'msal';

const isIE = window.navigator.userAgent.indexOf('MSIE ') > -1 || window.navigator.userAgent.indexOf('Trident/') > -1;

export const MsalConfig: Configuration = {
  auth: {
    clientId: '[YOUR-CLIENT-ID]',
    authority: 'https://login.microsoftonline.com/common/',
    validateAuthority: true,
    redirectUri: 'http://localhost:4200/',
    postLogoutRedirectUri: 'http://localhost:4200/',
    navigateToLoginRequestUrl: true
  },
  cache: {
    cacheLocation: 'localStorage',
    storeAuthStateInCookie: isIE // set to true for IE 11
  }
};

export const MSALAngularConfig: MsalAngularConfiguration = {
  popUp: !isIE,
  consentScopes: ['user.read', 'openid', 'profile']
};
