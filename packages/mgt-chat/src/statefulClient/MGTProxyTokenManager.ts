import { Providers } from '@microsoft/mgt-element';
export class MGTProxyTokenManager {
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

    this.proxyToken = await Providers.globalProvider.getAccessTokenForScopes({
      scopes: [Providers.globalProvider.webProxyAPIScope],
      forceTokenRefresh: true
    });

    await this.getToken(true);
    return this.proxyToken;
  };
}
