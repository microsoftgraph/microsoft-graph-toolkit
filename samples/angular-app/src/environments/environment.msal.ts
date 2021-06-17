// update these values with ones your AAD Client Id

import { MsalInterceptorConfiguration, MsalGuardConfiguration } from '@azure/msal-angular';
import { InteractionType, Configuration } from '@azure/msal-browser';

const isIE = window.navigator.userAgent.indexOf('MSIE ') > -1 || window.navigator.userAgent.indexOf('Trident/') > -1;

export const MsalConfig: Configuration = {
  auth: {
    clientId: 'a974dfa0-9f57-49b9-95db-90f04ce2111a',
    authority: 'https://login.microsoftonline.com/common/',
    redirectUri: 'http://localhost:4200/',
    postLogoutRedirectUri: 'http://localhost:4200/',
    navigateToLoginRequestUrl: true
  },
  cache: {
    cacheLocation: 'localStorage',
    storeAuthStateInCookie: isIE // set to true for IE 11
  }
};

const consentScopes = ['user.read', 'openid', 'profile'];
const protectedResourceMap = new Map<string, Array<string>>();
protectedResourceMap.set('https://graph.microsoft.com/v1.0/', consentScopes);


export const MsalInterceptorConfig: MsalInterceptorConfiguration = {
  interactionType: !isIE ? InteractionType.Popup : InteractionType.Redirect,
  protectedResourceMap
};

export const MsalGuardConfig: MsalGuardConfiguration = {
  interactionType: !isIE ? InteractionType.Popup : InteractionType.Redirect,
  authRequest: {
    scopes: consentScopes
  },
  loginFailedRoute: '/'
};
