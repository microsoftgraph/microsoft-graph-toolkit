import {
  AccountInfo,
  AuthenticationResult,
  AuthorizationCodeRequest,
  AuthorizationUrlRequest,
  Configuration,
  LogLevel,
  PublicClientApplication
} from '@azure/msal-node';
import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';
import { BrowserWindow, ipcMain } from 'electron';
import { cachePlugin } from './CachePlugin';
import { CustomFileProtocolListener } from './CustomFileProtocol';
import Main from './Main';

interface MsalElectronConfig {
  clientId: string;
  mainWindow: BrowserWindow;
  authority?: string;
  scopes?: string[];
}

export class ElectronAuthenticator {
  private ms_config: Configuration;
  public clientApplication: PublicClientApplication;
  public mainWindow: BrowserWindow;
  public authWindow: BrowserWindow;
  private account: AccountInfo;
  private authCodeUrlParams: AuthorizationUrlRequest;
  private authCodeRequest: AuthorizationCodeRequest;
  authCodeListener: CustomFileProtocolListener;
  silentRequest: { scopes: string[]; account: any; forceRefresh: boolean };

  constructor(config: MsalElectronConfig) {
    this.setConfig(config);
    this.clientApplication = new PublicClientApplication(this.ms_config);
    this.account = null;
    this.mainWindow = config.mainWindow;
    this.setupProvider();
    this.setRequestObjects();
  }

  //Set up all configuration variables
  private setConfig(config: MsalElectronConfig) {
    this.ms_config = {
      auth: {
        clientId: config.clientId,
        authority: config.authority
      },
      cache: {
        cachePlugin
      },
      system: {
        loggerOptions: {
          loggerCallback(loglevel, message, containsPii) {
            console.log(message);
          },
          piiLoggingEnabled: false,
          logLevel: LogLevel.Warning
        }
      }
    };
  }

  //Set up request parameters
  private setRequestObjects(scopes?): void {
    const requestScopes = scopes ? scopes : [];
    const redirectUri = 'msal://redirect';

    this.authCodeUrlParams = {
      scopes: requestScopes,
      redirectUri: redirectUri
    };

    this.authCodeRequest = {
      scopes: requestScopes,
      redirectUri: redirectUri,
      code: null
    };

    const baseSilentRequest = {
      account: null,
      forceRefresh: false
    };

    this.silentRequest = {
      ...baseSilentRequest,
      scopes: requestScopes
    };
  }

  //Set up an auth window with an option to be visible (invisible during silent sign in)
  setAuthWindow(visible: boolean) {
    this.authWindow = new BrowserWindow({ show: visible });
  }

  //set up messaging between authenticator and provider
  setupProvider() {
    this.mainWindow.webContents.on('did-finish-load', async () => {
      this.attemptSilentLogin();
    });

    ipcMain.handle('login', async () => {
      await this.login();
    });
    ipcMain.handle('token', async () => {
      const token = await this.getAccessToken();
      return token;
    });

    ipcMain.handle('logout', async () => {
      await this.logout();
    });

    ipcMain.on('tokenreceived', () => {
      console.log('Token received');
      if (!this.authWindow.isDestroyed()) {
        this.authWindow.close();
      }
    });
  }

  async getAccessToken(options?: AuthenticationProviderOptions): Promise<string> {
    let authResponse;
    const account = this.account || (await this.getAccount());
    if (this.authWindow.isDestroyed()) {
      this.setAuthWindow(false);
    }
    if (account) {
      this.silentRequest.account = account;
      authResponse = await this.getTokenSilent(this.silentRequest);
    } else {
      authResponse = await this.getTokenInteractive(false);
    }
    return authResponse.accessToken;
  }

  //Get token silently if available
  async getTokenSilent(tokenRequest): Promise<AuthenticationResult> {
    try {
      return await this.clientApplication.acquireTokenSilent(tokenRequest);
    } catch (error) {
      console.log('Silent token acquisition failed, acquiring token using redirect');
      return await this.getTokenInteractive(false);
    }
  }

  async login() {
    this.setAuthWindow(true);

    const authResponse = await this.getTokenInteractive(true);
    this.handleResponse(authResponse);
  }

  async logout(): Promise<void> {
    if (this.account) {
      await this.clientApplication.getTokenCache().removeAccount(this.account);
      this.account = null;
    }
  }

  //Set this.account
  private async handleResponse(response: AuthenticationResult) {
    if (response !== null) {
      this.account = response.account;
    } else {
      this.account = await this.getAccount();
    }

    return this.account;
  }

  //Get token interactively and optionally allow prompt to select account
  async getTokenInteractive(showPrompt: boolean): Promise<AuthenticationResult> {
    const authCodeUrlParams = showPrompt
      ? {
          ...this.authCodeUrlParams,
          scopes: this.authCodeUrlParams.scopes,
          prompt: 'select_account'
        }
      : {
          ...this.authCodeUrlParams,
          scopes: this.authCodeUrlParams.scopes
        };
    const authCodeUrl = await this.clientApplication.getAuthCodeUrl(authCodeUrlParams);
    this.authCodeListener = new CustomFileProtocolListener('msal');
    this.authCodeListener.start();
    const authCode = await this.listenForAuthCode(authCodeUrl);
    const authResult = await this.clientApplication.acquireTokenByCode({
      ...this.authCodeRequest,
      scopes: this.authCodeUrlParams.scopes,
      code: authCode
    });
    return authResult;
  }

  //Listen for the auth code in API response
  private async listenForAuthCode(navigateUrl: string): Promise<string> {
    this.authWindow.loadURL(navigateUrl);
    console.log('URL', navigateUrl);
    return new Promise((resolve, reject) => {
      this.authWindow.webContents.on('will-redirect', (event, responseUrl) => {
        try {
          console.log('Response URL', responseUrl);
          const parsedUrl = new URL(responseUrl);
          const authCode = parsedUrl.searchParams.get('code');
          resolve(authCode);
        } catch (err) {
          reject(err);
        }
      });
    });
  }

  //Attempt to Silently Sign In
  async attemptSilentLogin() {
    this.setAuthWindow(false);
    this.account = this.account || (await this.getAccount());
    if (this.account) {
      console.log('Account detected');
      this.mainWindow.webContents.send('isloggedin', true);
    } else {
      console.log('No accounts detected');
      this.mainWindow.webContents.send('isloggedin', false);
    }
  }

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
