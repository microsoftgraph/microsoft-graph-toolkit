import { Providers, log } from '@microsoft/mgt-element';
export class MGTProxyTokenManager {
  constructor() {
    this.refreshProxyToken();
  }

  private refreshProxyToken() {
    setInterval(
      () => {
        log('refreshing proxy token');
        this.tokenOperations().catch(error => {
          console.error('Error getting proxy token:', error);
        });
      },
      30 * 60 * 1000
    );
  }

  private readonly getToken = async (forceTokenRefresh = false) => {
    const token = await Providers.globalProvider.getAccessToken({ forceTokenRefresh });
    if (!token) throw new Error('Could not retrieve token for user');
    return token;
  };

  private proxyToken = '';
  public getProxyToken = async () => {
    if (this.proxyToken) {
      return this.proxyToken;
    }

    await this.tokenOperations();
    return this.proxyToken;
  };

  private async tokenOperations() {
    this.proxyToken = await Providers.globalProvider.getAccessTokenForScopes({
      scopes: [Providers.globalProvider.webProxyAPIScope],
      forceTokenRefresh: true
    });

    await this.getToken(true);
  }
}
