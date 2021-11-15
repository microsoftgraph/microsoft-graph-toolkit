/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

 import { IProvider, ProviderState, createFromProvider } from '@microsoft/mgt-element';
 import { loadConfiguration, TeamsUserCredential } from '@microsoft/teamsfx';
  
 /**
  * TeamsFx Provider handler
  *
  * @export
  * @class TeamsFxProvider
  * @extends {IProvider}
  */
 export class TeamsFxProvider extends IProvider {
   /**
    * returns _idToken
    *
    * @readonly
    * @type {boolean}
    * @memberof TeamsFxProvider
    */
   get isLoggedIn() {
     return !!this._idToken;
   }
 
   /**
    * Name used for analytics
    *
    * @readonly
    * @memberof IProvider
    */
   get name() {
     return 'MgtTeamsFxProvider';
   }
 
   /**
    * returns _credential
    *
    * @readonly
    * @memberof TeamsUserCredential
    */
   get credential() {
     return this._credential;
   }
 
   /**
    * privilege level for authentication
    *
    * @type {string[]}
    * @memberof SharePointProvider
    */
   scopes = [];
 
   _credential = null;
   /**
    * authority
    *
    * @type {string}
    * @memberof SharePointProvider
    */
   authority = '';
   _idToken = '';
 
   constructor(config, creds) {
     super();
     this.scopes = config.scopes;
 
     loadConfiguration({
        authentication: {
          clientId: config.clientId ?? process.env.REACT_APP_CLIENT_ID,
          initiateLoginEndpoint: config.initiateLoginEndpoint ?? process.env.REACT_APP_START_LOGIN_PAGE_URL,
          simpleAuthEndpoint: config.simpleAuthEndpoint ?? process.env.REACT_APP_TEAMSFX_ENDPOINT,
       },
     });
 
     if (creds) {
       this._credential = creds;
     } else {
       this._credential = new TeamsUserCredential();
     }
 
     this.graph = createFromProvider(this);
     this.internalLogin();
   }
 
   /**
    * uses provider to receive access token via SharePoint Provider
    *
    * @returns {Promise<string>}
    * @memberof SharePointProvider
    */
   async getAccessToken() {
     let accessToken = '';
     accessToken = (await this.credential.getToken(this.scopes)).token;
     return accessToken;
   }
   /**
    * update scopes
    *
    * @param {string[]} scopes
    * @memberof SharePointProvider
    */
   updateScopes(scopes) {
     this.scopes = scopes;
   }
 
  async internalLogin() {
    let token = null;

    try {
      token = await this.getAccessToken();
    } catch (error) {
      await this.credential.login(this.scopes);
    }
 
     if (!!token && token.expiresOnTimestamp < +new Date()) {
       await this.credential.login(this.scopes);
     }
 
     this._idToken = await this.getAccessToken();
     this.setState(this._idToken ? ProviderState.SignedIn : ProviderState.SignedOut);
   }
 }
 