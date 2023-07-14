/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { GraphPageIterator, IGraph, prepScopes } from '@microsoft/mgt-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export const getEventsQueryPageIterator = async (
  graph: IGraph,
  query: string,
  scopes = 'calendars.read'
): Promise<GraphPageIterator<MicrosoftGraph.Event>> => {
  const request = graph.api(query).middlewareOptions(prepScopes(scopes)).orderby('start/dateTime');

  return GraphPageIterator.create<MicrosoftGraph.Event>(graph, request);
};

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
export const getEventsPageIterator = async (
  graph: IGraph,
  startDateTime: Date,
  endDateTime: Date,
  groupId?: string
): Promise<GraphPageIterator<MicrosoftGraph.Event>> => {
  const sdt = `startdatetime=${startDateTime.toISOString()}`;
  const edt = `enddatetime=${endDateTime.toISOString()}`;

  let uri: string;

  if (groupId) {
    uri = `groups/${groupId}/calendar`;
  } else {
    uri = 'me';
  }

  uri += `/calendarview?${sdt}&${edt}`;

  return getEventsQueryPageIterator(graph, uri);
};
