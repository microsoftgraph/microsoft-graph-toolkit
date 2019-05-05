/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { IProvider, ProviderState } from '../providers/IProvider';
import { Graph } from '../Graph';
import { Client } from '@microsoft/microsoft-graph-client/lib/es/Client';

export class MockProvider extends IProvider {
  constructor(signedIn: boolean = false) {
    super();
    if (signedIn) {
      this.setState(ProviderState.SignedIn);
    } else {
      this.setState(ProviderState.SignedOut);
    }
  }

  async login(): Promise<void> {
    this.setState(ProviderState.Loading);
    await new Promise(resolve => setTimeout(resolve, 3000));
    this.setState(ProviderState.SignedIn);
  }

  async logout(): Promise<void> {
    this.setState(ProviderState.Loading);
    await new Promise(resolve => setTimeout(resolve, 3000));
    this.setState(ProviderState.SignedOut);
  }

  getAccessToken(): Promise<string> {
    return Promise.resolve('{token:https://graph.microsoft.com/}');
  }
  provider: any;

  graph = new MockGraph(this);
}

export class MockGraph extends Graph {
  private baseUrl = 'https://proxy.apisandbox.msdn.microsoft.com/svc?url=';
  private rootGraphUrl: string = 'https://graph.microsoft.com/';

  constructor(provider: MockProvider) {
    super(null);

    this.client = Client.initWithMiddleware({
      baseUrl: this.baseUrl + escape(this.rootGraphUrl),
      authProvider: provider
    });
  }

  async getEvents(startDateTime: Date, endDateTime: Date): Promise<Array<MicrosoftGraph.Event>> {
    let sdt = `startdatetime=${startDateTime.toISOString()}`;
    let edt = `enddatetime=${endDateTime.toISOString()}`;
    let uri = `/me/calendarview?${sdt}&${edt}`;

    let calendarView = await this.client.api(escape(uri)).get();
    return calendarView ? calendarView.value : null;
  }
}
