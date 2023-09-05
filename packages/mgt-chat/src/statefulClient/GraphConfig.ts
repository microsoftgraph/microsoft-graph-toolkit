import { GraphEndpoint } from '@microsoft/mgt-element';

export class GraphConfig {
  public static useCanary = false;

  public static version = 'testprodbetaembedforteamswebsockets';

  public static subscriptionConnectionVersion = 'testprodv1.0e2ewebsockets';

  public static get graphEndpoint(): GraphEndpoint {
    return GraphConfig.useCanary ? 'https://canary.graph.microsoft.com' : 'https://graph.microsoft.com';
  }

  private static get baseCanaryUrl() {
    return `https://canary.graph.microsoft.com/${GraphConfig.version}`;
  }

  public static get subscriptionEndpoint(): string {
    return GraphConfig.useCanary ? `${GraphConfig.baseCanaryUrl}/subscriptions` : '/subscriptions';
  }

  public static adjustNotificationUrl(url: string): string {
    let temp = url;
    if (GraphConfig.useCanary && url) {
      temp = url.replace('https://graph.microsoft.com/1.0', GraphConfig.baseCanaryUrl);
      temp = temp.replace('https://graph.microsoft.com/beta', GraphConfig.baseCanaryUrl);
    }
    temp = temp.replace(GraphConfig.version, GraphConfig.subscriptionConnectionVersion);
    return temp.replace('wss:', '');
  }
}
