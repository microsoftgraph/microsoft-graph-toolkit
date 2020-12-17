import { IProvider, Providers, ProviderState } from '@microsoft/mgt';
import { createFromProvider } from '@microsoft/mgt/dist/es6/Graph';
import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';
import { ipcRenderer } from 'electron';

export class ElectronProvider extends IProvider {
  constructor() {
    super();
    this.graph = createFromProvider(this);
    this.setupProvider();
  }

  setupProvider() {
    ipcRenderer.on('isloggedin', async (event, account) => {
      if (account) {
        Providers.globalProvider.setState(ProviderState.SignedIn);
      }
    });
  }
  async getAccessToken(
    options?: AuthenticationProviderOptions
  ): Promise<string> {
    const token = await ipcRenderer.invoke('token');
    console.log('Token is ', token);
    ipcRenderer.send('tokenreceived');
    return token;
  }
  async login(): Promise<void> {
    await ipcRenderer.invoke('login');
    Providers.globalProvider.setState(ProviderState.SignedIn);
  }

  async logout(): Promise<void> {
    await ipcRenderer.invoke('logout');
    Providers.globalProvider.setState(ProviderState.SignedOut);
  }
}
