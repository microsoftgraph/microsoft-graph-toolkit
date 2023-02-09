/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  AccountInfo,
  AuthenticationResult,
  AuthError,
  AuthorizationCodeRequest,
  AuthorizationUrlRequest,
  Configuration,
  ICachePlugin,
  LogLevel,
  PublicClientApplication
} from '@azure/msal-node';
import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client';
import { GraphEndpoint } from '@microsoft/mgt-element';
import { BrowserWindow, ipcMain } from 'electron';
import { CustomFileProtocolListener } from './CustomFileProtocol';
import { REDIRECT_URI, COMMON_AUTHORITY_URL } from './Constants';
/**
 * base config for MSAL authentication
 *
 * @interface MsalElectronConfig
 */
export interface MsalElectronConfig {
  /**
   * Client ID alphanumeric code
   *
   * @type {string}
   * @memberof MsalElectronConfig
   */
  clientId: string;

  /**
   * Main window instance
   *
   * @type {BrowserWindow}
   * @memberof MsalElectronConfig
   */
  mainWindow: BrowserWindow;

  /**
   * Config authority
   *
   * @type {string}
   * @memberof MsalElectronConfig
   */
  authority?: string;

  /**
   * List of scopes
   *
   * @type {string[]}
   * @memberof MsalElectronConfig
   */
  scopes?: string[];

  /**
   * Cache plugin to enable persistent caching
   *
   * @type {ICachePlugin}
   * @memberof MsalElectronConfig
   */
  cachePlugin?: ICachePlugin;
  /**
   * The base URL for the graph client
   */
  baseURL?: GraphEndpoint;
}

/**
 * Prompt type for consent or login
 *
 * @enum {number}
 */
enum promptType {
  SELECT_ACCOUNT = 'select_account'
}

/**
 * State of Authentication Provider
 *
 * @enum {number}
 */
enum AuthState {
  LOGGED_IN = 'logged_in',
  LOGGED_OUT = 'logged_out'
}

/**
 * AccountDetails defines the available AccountInfo or undefined.
 */
type AccountDetails = AccountInfo | undefined;

/**
 * ElectronAuthenticator class to be instantiated in the main process.
 * Responsible for MSAL authentication flow and token acqusition.
 *
 * @export
 * @class ElectronAuthenticator
 */
export class ElectronAuthenticator {
  /**
   * Configuration for MSAL Authentication
   *
   * @private
   * @type {Configuration}
   * @memberof ElectronAuthenticator
   */
  private ms_config: Configuration;

  /**
   * Application instance
   *
   * @type {PublicClientApplication}
   * @memberof ElectronAuthenticator
   */
  public clientApplication: PublicClientApplication;

  /**
   * Mainwindow instance
   *
   * @type {BrowserWindow}
   * @memberof ElectronAuthenticator
   */
  public mainWindow: BrowserWindow;

  // Popup which will take the user through the login/consent process

  /**
   *
   *
   * @type {BrowserWindow}
   * @memberof ElectronAuthenticator
   */
  public authWindow: BrowserWindow;

  /**
   * Logged in account
   *
   * @private
   * @type {AccountDetails}
   * @memberof ElectronAuthenticator
   */
  private account: AccountDetails;

  /**
   * Params to generate the URL for MSAL auth
   *
   * @private
   * @type {AuthorizationUrlRequest}
   * @memberof ElectronAuthenticator
   */
  private authCodeUrlParams: AuthorizationUrlRequest;

  /**
   * Request for authentication call
   *
   * @private
   * @type {AuthorizationCodeRequest}
   * @memberof ElectronAuthenticator
   */
  private authCodeRequest: AuthorizationCodeRequest;

  /**
   * Listener that will listen for auth code in response
   *
   * @private
   * @type {CustomFileProtocolListener}
   * @memberof ElectronAuthenticator
   */
  private authCodeListener: CustomFileProtocolListener;

  /**
   * Instance of the authenticator
   *
   * @private
   * @static
   * @type {ElectronAuthenticator}
   * @memberof ElectronAuthenticator
   */
  private static authInstance: ElectronAuthenticator;

  /**
   * Creates an instance of ElectronAuthenticator.
   * @param {MsalElectronConfig} config
   * @memberof ElectronAuthenticator
   */
  private constructor(config: MsalElectronConfig) {
    this.setConfig(config);
    this.account = undefined;
    this.mainWindow = config.mainWindow;
    this.setRequestObjects(config.scopes);
    this.setupProvider();
  }

  /**
   * Initialize the authenticator. Call this method in your main process to create an instance of ElectronAuthenticator.
   *
   * @static
   * @param {MsalElectronConfig} config
   * @memberof ElectronAuthenticator
   */
  public static initialize(config: MsalElectronConfig) {
    if (!ElectronAuthenticator.instance) {
      ElectronAuthenticator.authInstance = new ElectronAuthenticator(config);
    }
  }

  /**
   * Getter for the ElectronAuthenticator instance.
   *
   * @readonly
   * @memberof ElectronAuthenticator
   */
  public static get instance() {
    return this.authInstance;
  }

  /**
   * Setting up config for MSAL auth
   *
   * @private
   * @param {MsalElectronConfig} config
   * @memberof ElectronAuthenticator
   */
  private async setConfig(config: MsalElectronConfig) {
    this.ms_config = {
      auth: {
        clientId: config.clientId,
        authority: config.authority ? config.authority : COMMON_AUTHORITY_URL
      },
      cache: config.cachePlugin ? { cachePlugin: config.cachePlugin } : undefined,
      system: {
        loggerOptions: {
          loggerCallback(loglevel, message, containsPii) {},
          piiLoggingEnabled: false,
          logLevel: LogLevel.Warning
        }
      }
    };
    this.clientApplication = new PublicClientApplication(this.ms_config);
  }

  /**
   * Set up request parameters
   *
   * @protected
   * @param {*} [scopes]
   * @memberof ElectronAuthenticator
   */
  protected setRequestObjects(scopes?): void {
    const requestScopes = scopes ? scopes : [];
    const redirectUri = REDIRECT_URI;

    this.authCodeUrlParams = {
      scopes: requestScopes,
      redirectUri
    };

    this.authCodeRequest = {
      scopes: requestScopes,
      redirectUri,
      code: ''
    };
  }

  /**
   * Set up an auth window with an option to be visible (invisible during silent sign in)
   *
   * @protected
   * @param {boolean} visible
   * @memberof ElectronAuthenticator
   */
  protected setAuthWindow(visible: boolean) {
    this.authWindow = new BrowserWindow({ show: visible });
  }

  /**
   * Set up messaging between authenticator and provider
   * @protected
   * @memberof ElectronAuthenticator
   */
  protected setupProvider() {
    this.mainWindow.webContents.on('did-finish-load', async () => {
      await this.attemptSilentLogin();
    });

    ipcMain.handle('login', async () => {
      const account = await this.login();
      if (account) {
        this.mainWindow.webContents.send('mgtAuthState', AuthState.LOGGED_IN);
      } else {
        this.mainWindow.webContents.send('mgtAuthState', AuthState.LOGGED_OUT);
      }
    });
    ipcMain.handle('token', async (e, options: AuthenticationProviderOptions) => {
      try {
        const token = await this.getAccessToken(options);
        return token;
      } catch (e) {
        throw e;
      }
    });

    ipcMain.handle('logout', async () => {
      await this.logout();
      this.mainWindow.webContents.send('mgtAuthState', AuthState.LOGGED_OUT);
    });
  }

  /**
   * Get access token
   *
   * @protected
   * @param {AuthenticationProviderOptions} [options]
   * @return {*}  {Promise<string>}
   * @memberof ElectronAuthenticator
   */
  protected async getAccessToken(options?: AuthenticationProviderOptions): Promise<string | undefined> {
    let authResponse;
    const scopes = options && options.scopes ? options.scopes : this.authCodeUrlParams.scopes;
    const account = this.account || (await this.getAccount());
    if (account) {
      const request = {
        account,
        scopes,
        forceRefresh: false
      };
      authResponse = await this.getTokenSilent(request, scopes);
    }
    if (authResponse) {
      return authResponse.accessToken;
    }
    return undefined;
  }

  /**
   * Get token silently if available
   *
   * @protected
   * @param {*} tokenRequest
   * @param {*} [scopes]
   * @return {*}  {Promise<AuthenticationResult>}
   * @memberof ElectronAuthenticator
   */
  protected async getTokenSilent(tokenRequest, scopes?): Promise<AuthenticationResult | null> {
    try {
      return await this.clientApplication.acquireTokenSilent(tokenRequest);
    } catch (error) {
      return null;
    }
  }

  /**
   * Login (open popup and allow user to select account/login)
   * @private
   * @return {*}
   * @memberof ElectronAuthenticator
   */
  protected async login() {
    const authResponse = await this.getTokenInteractive(promptType.SELECT_ACCOUNT);
    return this.setAccountFromResponse(authResponse);
  }

  /**
   * Logout
   *
   * @private
   * @return {*}  {Promise<void>}
   * @memberof ElectronAuthenticator
   */
  protected async logout(): Promise<void> {
    if (this.account) {
      await this.clientApplication.getTokenCache().removeAccount(this.account);
      this.account = undefined;
    }
  }

  /**
   * Set this.account to current logged in account
   *
   * @private
   * @param {AuthenticationResult} response
   * @return {*}
   * @memberof ElectronAuthenticator
   */
  private async setAccountFromResponse(response: AuthenticationResult) {
    if (response) {
      this.account = response?.account || undefined;
    } else {
      this.account = await this.getAccount();
    }
    return this.account;
  }

  /**
   * Get token interactively and optionally allow prompt to select account
   *
   * @protected
   * @param {promptType} prompt_type
   * @param {*} [scopes]
   * @return {*}  {Promise<AuthenticationResult>}
   * @memberof ElectronAuthenticator
   */
  protected async getTokenInteractive(prompt_type: promptType, scopes?): Promise<AuthenticationResult> {
    let authResult;
    const requestScopes = scopes ? scopes : this.authCodeUrlParams.scopes;
    const authCodeUrlParams = {
      ...this.authCodeUrlParams,
      scopes: requestScopes,
      prompt: prompt_type.toString()
    };
    const authCodeUrl = await this.clientApplication.getAuthCodeUrl(authCodeUrlParams);
    this.authCodeListener = new CustomFileProtocolListener('msal');
    this.authCodeListener.start();
    const authCode = await this.listenForAuthCode(authCodeUrl, prompt_type);
    authResult = await this.clientApplication
      .acquireTokenByCode({
        ...this.authCodeRequest,
        scopes: requestScopes,
        code: authCode || ''
      })
      .catch((e: AuthError) => {
        throw e;
      });
    return authResult;
  }

  /**
   * Listen for the auth code in API response
   *
   * @private
   * @param {string} navigateUrl
   * @param {promptType} prompt_type
   * @return {*}  {Promise<string>}
   * @memberof ElectronAuthenticator
   */
  private async listenForAuthCode(navigateUrl: string, prompt_type: promptType): Promise<string | null> {
    this.setAuthWindow(true);
    await this.authWindow.loadURL(navigateUrl);
    return new Promise((resolve, reject) => {
      this.authWindow.webContents.on('will-redirect', (event, responseUrl) => {
        try {
          const parsedUrl = new URL(responseUrl);
          const authCode = parsedUrl.searchParams.get('code');
          resolve(authCode);
        } catch (err) {
          this.authWindow.destroy();
          reject(err);
        }
        this.authWindow.destroy();
      });
    });
  }

  /**
   * Attempt to Silently Sign In
   *
   * @protected
   * @memberof ElectronAuthenticator
   */
  protected async attemptSilentLogin() {
    this.account = this.account || (await this.getAccount());
    if (this.account) {
      const token = await this.getAccessToken();
      if (token) {
        this.mainWindow.webContents.send('mgtAuthState', AuthState.LOGGED_IN);
      } else {
        this.mainWindow.webContents.send('mgtAuthState', AuthState.LOGGED_OUT);
      }
    } else {
      this.mainWindow.webContents.send('mgtAuthState', AuthState.LOGGED_OUT);
    }
  }

  /**
   * Get logged in Account details
   *
   * @private
   * @return {*}  {Promise<AccountInfo>}
   * @memberof ElectronAuthenticator
   */
  private async getAccount(): Promise<AccountDetails> {
    const cache = this.clientApplication.getTokenCache();
    const currentAccounts = await cache.getAllAccounts();
    if (currentAccounts?.length >= 1) {
      return currentAccounts[0];
    }
    return undefined;
  }
}
