import { GraphEndpoint } from '@microsoft/mgt-element';

export class GraphConfig {
  public static useCanary = false;

  public static version = 'v1.0';

  public static canarySubscriptionVersion = 'testprodv1.0e2ewebsockets';

  public static webSocketsPrefix = 'websockets:';

  public static get graphEndpoint(): GraphEndpoint {
    return GraphConfig.useCanary ? 'https://canary.graph.microsoft.com' : 'https://graph.microsoft.com';
  }

  public static get baseCanaryUrl(): string {
    return `${GraphConfig.graphEndpoint}/${this.canarySubscriptionVersion}`;
  }

  public static get subscriptionEndpoint(): string {
    return GraphConfig.useCanary ? `${GraphConfig.baseCanaryUrl}/subscriptions` : '/subscriptions';
  }

  public static adjustNotificationUrl(url: string): string {
    if (GraphConfig.useCanary && url) {
      url = url.replace('https://graph.microsoft.com/1.0', GraphConfig.baseCanaryUrl);
      url = url.replace('https://graph.microsoft.com/beta', GraphConfig.baseCanaryUrl);
    }
    return url.replace(GraphConfig.webSocketsPrefix, '');
  }
}
