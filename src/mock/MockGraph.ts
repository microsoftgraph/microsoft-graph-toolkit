/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Client } from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { BaseGraph } from '../BaseGraph';
import { MgtBaseComponent } from '../components/baseComponent';
import { MockProvider } from './MockProvider';

/**
 * MockGraph Instance
 *
 * @export
 * @class MockGraph
 * @extends {Graph}
 */
// tslint:disable-next-line: max-classes-per-file
export class MockGraph extends BaseGraph {
  private static readonly BASE_URL = 'https://proxy.apisandbox.msdn.microsoft.com/svc?url=';
  private static readonly ROOT_GRAPH_URL: string = 'https://graph.microsoft.com/';

  constructor(mockProvider: MockProvider) {
    super(mockProvider);

    this.client = Client.initWithMiddleware({
      authProvider: mockProvider,
      baseUrl: MockGraph.BASE_URL + escape(MockGraph.ROOT_GRAPH_URL)
    });
  }

  /**
   * Returns an instance of the Graph in the context of the provided component.
   *
   * @param {MgtBaseComponent} component
   * @returns
   * @memberof Graph
   */
  public forComponent(component: MgtBaseComponent): MockGraph {
    // The purpose of the forComponent pattern is to update the headers of any outgoing Graph requests.
    // The MockGraph isn't making real Graph requests, so we can simply no-op and return the same instance.
    return this;
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
