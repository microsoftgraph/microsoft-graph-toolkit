/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { GraphEndpoint } from '@microsoft/mgt-element';
import { replaceOrAppendSessionId } from './replaceOrAppendSessionId';

export class GraphConfig {
  public static ackAsString = false;

  public static useCanary = false;

  public static version = 'v1.0';

  public static canarySubscriptionVersion = 'testprodv1.0e2ewebsockets';

  public static webSocketsPrefix = 'websockets:';
  static usePremiumApis = false;

  public static get graphEndpoint(): GraphEndpoint {
    return GraphConfig.useCanary ? 'https://canary.graph.microsoft.com' : 'https://graph.microsoft.com';
  }

  public static get baseCanaryUrl(): string {
    return `${GraphConfig.graphEndpoint}/${this.canarySubscriptionVersion}`;
  }

  public static get subscriptionEndpoint(): string {
    return GraphConfig.useCanary ? `${GraphConfig.baseCanaryUrl}/subscriptions` : '/subscriptions';
  }

  public static adjustNotificationUrl(url: string, sessionId = 'default'): string {
    if (GraphConfig.useCanary && url) {
      url = url.replace('https://graph.microsoft.com/1.0', GraphConfig.baseCanaryUrl);
      url = url.replace('https://graph.microsoft.com/beta', GraphConfig.baseCanaryUrl);
    }
    url = url.replace(GraphConfig.webSocketsPrefix, '');
    // update or append sessionid
    url = replaceOrAppendSessionId(url, sessionId);
    return url;
  }
}
