/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Client } from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { Graph } from '../Graph';
import { IProvider, ProviderState } from '../providers/IProvider';
/**
 * Mock Provider access token for Microsoft Graph APIs
 *
 * @export
 * @class MockProvider
 * @extends {IProvider}
 */
export class MockProvider extends IProvider {
  // tslint:disable-next-line: completed-docs
  public provider: any;

  /**
   * new instance of mock graph provider
   *
   * @memberof MockProvider
   */
  public graph = new MockGraph(this);
  constructor(signedIn: boolean = false) {
    super();
    if (signedIn) {
      this.setState(ProviderState.SignedIn);
    } else {
      this.setState(ProviderState.SignedOut);
    }
  }

  /**
   * sets Provider state to SignedIn
   *
   * @returns {Promise<void>}
   * @memberof MockProvider
   */
  public async login(): Promise<void> {
    this.setState(ProviderState.Loading);
    await new Promise(resolve => setTimeout(resolve, 3000));
    this.setState(ProviderState.SignedIn);
  }
  /**
   * sets Provider state to signed out
   *
   * @returns {Promise<void>}
   * @memberof MockProvider
   */
  public async logout(): Promise<void> {
    this.setState(ProviderState.Loading);
    await new Promise(resolve => setTimeout(resolve, 3000));
    this.setState(ProviderState.SignedOut);
  }
  /**
   * Promise returning token from graph.microsoft.com
   *
   * @returns {Promise<string>}
   * @memberof MockProvider
   */
  public getAccessToken(): Promise<string> {
    return Promise.resolve('{token:https://graph.microsoft.com/}');
  }
}

/**
 * MockGraph Instance
 *
 * @export
 * @class MockGraph
 * @extends {Graph}
 */
// tslint:disable-next-line: max-classes-per-file
export class MockGraph extends Graph {
  private baseUrl = 'https://proxy.apisandbox.msdn.microsoft.com/svc?url=';
  private rootGraphUrl: string = 'https://graph.microsoft.com/';

  constructor(provider: MockProvider) {
    super(null);

    this.client = Client.initWithMiddleware({
      authProvider: provider,
      baseUrl: this.baseUrl + escape(this.rootGraphUrl)
    });
  }
/**
 * get events for Calendar
 *
 * @param {Date} startDateTime
 * @param {Date} endDateTime
 * @returns {Promise<MicrosoftGraph.Event[]>}
 * @memberof MockGraph
 */
public async getEvents(startDateTime: Date, endDateTime: Date): Promise<MicrosoftGraph.Event[]> {
    const sdt = `startdatetime=${startDateTime.toISOString()}`;
    const edt = `enddatetime=${endDateTime.toISOString()}`;
    const uri = `/me/calendarview?${sdt}&${edt}`;

    const calendarView = await this.client.api(escape(uri)).get();
    return calendarView ? calendarView.value : null;
  }
}
