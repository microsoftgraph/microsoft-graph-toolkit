import { GraphEndpoint } from '@microsoft/mgt-element';

export class GraphConfig {
  public static useCanary = false;
  public static get graphEndpoint(): GraphEndpoint {
    return GraphConfig.useCanary ? 'https://canary.graph.microsoft.com' : 'https://graph.microsoft.com';
  }
  private static get baseCanaryUrl() {
    return 'https://canary.graph.microsoft.com/testprodv1.0e2ewebsockets';
  }
  public static get subscriptionEndpoint(): string {
    return GraphConfig.useCanary ? `${GraphConfig.baseCanaryUrl}/subscriptions` : '/subscriptions';
  }
  public static adjustNotificationUrl(url: string): string {
    let temp = url;
    if (GraphConfig.useCanary && url) {
      temp = url.replace('https://graph.microsoft.com/1.0', GraphConfig.baseCanaryUrl);
    }
    return temp.replace('wss:', '');
  }
}
