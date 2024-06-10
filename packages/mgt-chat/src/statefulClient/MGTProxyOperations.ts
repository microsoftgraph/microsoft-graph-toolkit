import axios from 'axios';
import type { Subscription } from '@microsoft/microsoft-graph-types';

export interface NegotiateData {
  negotiateVersion: number | undefined;
  url: string | undefined;
  accessToken: string | undefined;
}

export interface ProxySubscription {
  subscription: Subscription | undefined;
  negotiate: NegotiateData | undefined;
}

export interface RenewedProxySubscription {
  subscription: Subscription | undefined;
}

export class MGTProxyOperations {
  public static async PerformOperation(
    url: string,
    method: string,
    operationData: Subscription,
    accesstoken: string
  ): Promise<ProxySubscription | undefined> {
    try {
      const response = await axios({
        url,
        method,
        headers: {
          'Content-Type': 'application/json',
          Authorization: 'Bearer ' + accesstoken
        },
        data: operationData
      });
      const data = response.data as ProxySubscription;
      return data;
    } catch (error) {
      console.error('Error:', error);
    }
  }
}
