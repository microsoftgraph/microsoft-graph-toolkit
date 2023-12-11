/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { GraphPageIterator, IGraph, prepScopes } from '@microsoft/mgt-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

/**
 *
 * @param {IGraph} graph
 * @param {string} query the graph resource and query string to be requested
 * @param {string[]} additionalScopes an array of scope to be requested before making the request
 * Should be calculated by the calling code using `IProvider.needsAdditionalScopes()`
 * @returns {Promise<GraphPageIterator<MicrosoftGraph.Event>>} a page iterator to allow
 * the calling code to request more data if present and needed
 */
export const getEventsQueryPageIterator = async (
  graph: IGraph,
  query: string,
  additionalScopes: string[]
): Promise<GraphPageIterator<MicrosoftGraph.Event>> => {
  const request = graph.api(query).middlewareOptions(prepScopes(additionalScopes)).orderby('start/dateTime');

  return GraphPageIterator.create<MicrosoftGraph.Event>(graph, request);
};

/**
 * returns Calender events iterator associated with either the logged in user or a specific groupId
 *
 * @param {IGraph} graph
 * @param {Date} startDateTime
 * @param {Date} endDateTime
 * @param {string} [groupId]
 * @returns {Promise<GraphPageIterator<MicrosoftGraph.Event>>}
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

  const uri: string = groupId
    ? `groups/${groupId}/calendar/calendarview?${sdt}&${edt}`
    : `me/calendarview?${sdt}&${edt}`;
  const allValidScopes = groupId
    ? ['Group.Read.All', 'Group.ReadWrite.All']
    : ['Calendars.ReadBasic', 'Calendars.Read', 'Calendars.ReadWrite'];

  return getEventsQueryPageIterator(graph, uri, allValidScopes);
};
