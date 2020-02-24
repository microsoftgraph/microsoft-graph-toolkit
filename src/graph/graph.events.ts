/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { IGraph } from '../IGraph';
import { prepScopes } from '../utils/GraphHelpers';

/**
 * async promise, returns Calender events associated with either the logged in user or a specific groupId
 *
 * @param {Date} startDateTime
 * @param {Date} endDateTime
 * @param {string} [groupId]
 * @returns {(Promise<Event[]>)}
 * @memberof Graph
 */
export async function getEvents(
  graph: IGraph,
  startDateTime: Date,
  endDateTime: Date,
  groupId?: string
): Promise<MicrosoftGraph.Event[]> {
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

  const calendarView = await graph
    .api(uri)
    .middlewareOptions(prepScopes(scopes))
    .orderby('start/dateTime')
    .get();
  return calendarView ? calendarView.value : null;
}
