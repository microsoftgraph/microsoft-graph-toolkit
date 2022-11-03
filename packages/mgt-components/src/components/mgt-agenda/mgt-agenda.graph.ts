/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { GraphPageIterator, IGraph, prepScopes } from '@microsoft/mgt-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

/**
 * returns Calender events iterator associated with either the logged in user or a specific groupId
 *
 * @param {Date} startDateTime
 * @param {Date} endDateTime
 * @param {string} [groupId]
 * @param {string} preferredTimezone
 * @returns {(Promise<Event[]>)}
 * @memberof Graph
 */
export async function getEventsPageIterator(
  graph: IGraph,
  startDateTime: Date,
  endDateTime: Date,
  groupId?: string
): Promise<GraphPageIterator<MicrosoftGraph.Event>> {
  const scopes = 'calendars.read';

  const sdt = `startdatetime=${startDateTime.toISOString()}`;
  const edt = `enddatetime=${endDateTime.toISOString()}`;

  let uri: string;

  if (groupId) {
    uri = `groups/${groupId}/calendar`;
  } else {
    uri = 'me';
  }

  uri += `/calendarview?${sdt}&${edt}`;

  let request = graph.api(uri).middlewareOptions(prepScopes(scopes)).orderby('start/dateTime');

  return GraphPageIterator.create<MicrosoftGraph.Event>(graph, request);
}
