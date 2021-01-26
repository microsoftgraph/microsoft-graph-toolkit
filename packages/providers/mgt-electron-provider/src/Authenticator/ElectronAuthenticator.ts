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
import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';
import { BrowserWindow, ipcMain } from 'electron';
import { CustomFileProtocolListener } from './CustomFileProtocol';
import { REDIRECT_URI } from './Constants';

/**
 * base config for MSAL authentication
 *
 * @interface MsalElectronConfig
 */
interface MsalElectronConfig {
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
   * Incremental consent, false by default
   *
   * @type {boolean}
   * @memberof MsalElectronConfig
   */
  isIncrementalConsentEnabled?: boolean;

  cachePlugin?: ICachePlugin;
}

/**
 * Prompt type for consent or login
 *
 * @enum {number}
 */
enum Prompt_Type {
  select_account = 'select_account',
  consent = 'consent'
}

/**
 * ElectronAuthenticator class to be instantiated in the main process.
 * Responsible for MSAL authentication flow and token acqusition.
 *
 * @export
 * @class ElectronAuthenticator
 */
export class ElectronAuthenticator {
  //config variables to set up MSAL auth
  private ms_config: Configuration;

  //Enable incremental consent, false by default
  private isIncrementalConsentEnabled: Boolean = false;

  //Application instance
  public clientApplication: PublicClientApplication;

  //Mainwindow instance
  public mainWindow: BrowserWindow;

  //Popup which will take the user through the login/consent process
  public authWindow: BrowserWindow;

  //Logged in account
  private account: AccountInfo;

  //Params to generate the URL for MSAL auth
  private authCodeUrlParams: AuthorizationUrlRequest;

  //Request for authentication call
  private authCodeRequest: AuthorizationCodeRequest;

  //Listener that will listen for auth code in response
  authCodeListener: CustomFileProtocolListener;

  //List of scopes denied by user so that they are not asked again
  private deniedScopes: string[];

  //List of scopes that are initially consented
  scopes: string[];

  constructor(config: MsalElectronConfig) {
    this.setConfig(config);
    this.account = null;
    this.mainWindow = config.mainWindow;
    this.setRequestObjects(config.scopes);
    this.isIncrementalConsentEnabled = config.isIncrementalConsentEnabled ? true : false;
    this.setupProvider();
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
        authority: config.authority
      },
      cache: config.cachePlugin ? { cachePlugin: config.cachePlugin } : null,
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
   * @private
   * @param {*} [scopes]
   * @memberof ElectronAuthenticator
   */
  private setRequestObjects(scopes?): void {
    const requestScopes = scopes ? scopes : [];
    const redirectUri = REDIRECT_URI;

    this.authCodeUrlParams = {
      scopes: requestScopes,
      redirectUri: redirectUri
    };

    this.authCodeRequest = {
      scopes: requestScopes,
      redirectUri: redirectUri,
      code: null
    };
  }

  /**
   * Set up an auth window with an option to be visible (invisible during silent sign in)
   *
   * @param {boolean} visible
   * @memberof ElectronAuthenticator
   */
  setAuthWindow(visible: boolean) {
    this.authWindow = new BrowserWindow({ show: visible });
  }

  /**
   * Set up messaging between authenticator and provider
   *
   * @memberof ElectronAuthenticator
   */
  setupProvider() {
    this.mainWindow.webContents.on('did-finish-load', async () => {
      await this.attemptSilentLogin();
    });

    ipcMain.handle('login', async () => {
      const account = await this.login();
      if (account) {
        this.mainWindow.webContents.send('isloggedin', true);
      } else {
        this.mainWindow.webContents.send('isloggedin', false);
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
      this.mainWindow.webContents.send('isloggedin', false);
    });
  }

  /**
   * Get access token
   *
   * @param {AuthenticationProviderOptions} [options]
   * @return {*}  {Promise<string>}
   * @memberof ElectronAuthenticator
   */
  async getAccessToken(options?: AuthenticationProviderOptions): Promise<string> {
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
    } else {
      throw new Error('Permission Denied');
    }
  }

  /**
   * Get token silently if available
   *
   * @param {*} tokenRequest
   * @param {*} [scopes]
   * @return {*}  {Promise<AuthenticationResult>}
   * @memberof ElectronAuthenticator
   */
  async getTokenSilent(tokenRequest, scopes?): Promise<AuthenticationResult> {
    let token;
    try {
      return await this.clientApplication.acquireTokenSilent(tokenRequest);
    } catch (error) {
      if (!this.isIncrementalConsentEnabled) {
        return null;
      }
      console.log('Silent token acquisition failed, acquiring token using redirect');
      if (!this.areScopesDenied(scopes)) {
        token = await this.getTokenInteractive(Prompt_Type.consent, scopes);
      } else {
        throw new Error('Scopes are denied');
      }
      return token;
    }
  }

  /**
   * Login (open popup and allow user to select account/login)
   *
   * @return {*}
   * @memberof ElectronAuthenticator
   */
  async login() {
    const authResponse = await this.getTokenInteractive(Prompt_Type.select_account);
    return this.setAccountFromResponse(authResponse);
  }

  /**
   * Logout
   *
   * @return {*}  {Promise<void>}
   * @memberof ElectronAuthenticator
   */
  async logout(): Promise<void> {
    if (this.account) {
      await this.clientApplication.getTokenCache().removeAccount(this.account);
      this.account = null;
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
    if (response !== null) {
      this.account = response.account;
    } else {
      this.account = await this.getAccount();
    }
    return this.account;
  }

  /**
   * Get token interactively and optionally allow prompt to select account
   *
   * @param {Prompt_Type} prompt_type
   * @param {*} [scopes]
   * @return {*}  {Promise<AuthenticationResult>}
   * @memberof ElectronAuthenticator
   */
  async getTokenInteractive(prompt_type: Prompt_Type, scopes?): Promise<AuthenticationResult> {
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
        code: authCode
      })
      .catch((e: AuthError) => {
        console.log('Denied scopes');
        this.addDeniedScopes(requestScopes);
      });
    return authResult;
  }

  /**
   * Check if scopes have been previously denied
   *
   * @protected
   * @param {string[]} scopes
   * @return {*}
   * @memberof ElectronAuthenticator
   */
  protected areScopesDenied(scopes: string[]) {
    if (scopes) {
      if (this.deniedScopes && this.deniedScopes.filter(s => -1 !== scopes.indexOf(s)).length > 0) {
        return true;
      }
    }
    return false;
  }

  /**
   * If scopes are denied, add them to the list of denied scopes
   *
   * @protected
   * @param {string[]} scopes
   * @memberof ElectronAuthenticator
   */
  protected addDeniedScopes(scopes: string[]) {
    if (scopes) {
      let deniedScopes: string[] = this.deniedScopes || [];
      this.deniedScopes = deniedScopes.concat(scopes);

      let index = deniedScopes.indexOf('openid');
      if (index !== -1) {
        deniedScopes.splice(index, 1);
      }

      index = deniedScopes.indexOf('profile');
      if (index !== -1) {
        deniedScopes.splice(index, 1);
      }
    }
  }

  /**
   * Listen for the auth code in API response
   *
   * @private
   * @param {string} navigateUrl
   * @param {Prompt_Type} prompt_type
   * @return {*}  {Promise<string>}
   * @memberof ElectronAuthenticator
   */
  private async listenForAuthCode(navigateUrl: string, prompt_type: Prompt_Type): Promise<string> {
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
   * @memberof ElectronAuthenticator
   */
  async attemptSilentLogin() {
    this.account = this.account || (await this.getAccount());
    if (this.account) {
      const token = await this.getAccessToken();
      if (token) {
        this.mainWindow.webContents.send('isloggedin', true);
      } else {
        console.log('No token');
        this.mainWindow.webContents.send('isloggedin', false);
      }
    } else {
      console.log('No accounts detected');
      this.mainWindow.webContents.send('isloggedin', false);
    }
  }

  /**
   * Get logged in Account details
   *
   * @private
   * @return {*}  {Promise<AccountInfo>}
   * @memberof ElectronAuthenticator
   */
  private async getAccount(): Promise<AccountInfo> {
    const cache = this.clientApplication.getTokenCache();
    const currentAccounts = await cache.getAllAccounts();

    if (currentAccounts === null) {
      console.log('No accounts detected');
      return null;
    }

    if (currentAccounts.length > 1) {
      // Add choose account code here
      console.log('Multiple accounts detected, need to add choose account code.');
      return currentAccounts[0];
    } else if (currentAccounts.length === 1) {
      return currentAccounts[0];
    } else {
      return null;
    }
  }
}
