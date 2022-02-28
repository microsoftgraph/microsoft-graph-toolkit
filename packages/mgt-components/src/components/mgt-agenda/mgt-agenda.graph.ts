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
  groupId?: string,
  preferredTimezone?: string
): Promise<GraphPageIterator<MicrosoftGraph.Event>> {
  const scopes = 'calendars.read';

  let sdt = `startdatetime=${dateToLocalISO(startDateTime)}`;
  let edt = `enddatetime=${dateToLocalISO(endDateTime)}`;

  let uri: string;

  if (groupId) {
    uri = `groups/${groupId}/calendar`;
  } else {
    uri = 'me';
  }

  uri += `/calendarview?${sdt}&${edt}`;

  let request = graph.api(uri).middlewareOptions(prepScopes(scopes)).orderby('start/dateTime');

  if (preferredTimezone) {
    request = request.header('Prefer', `outlook.timezone="${preferredTimezone}"`);
  }

  return GraphPageIterator.create<MicrosoftGraph.Event>(graph, request);
}

/**
 * Convert a date object to a local time ISO string.
 * @param date Date object.
 * @returns ISO 8601 string with timezone offset.
 */
function dateToLocalISO(date: Date): string {
  // Get difference, in minutes, between date, as evaluated in the
  // UTC time zone, and as evaluated in the local time zone.
  // Why? The values of startDateTime and endDateTime are interpreted
  // using the timezone offset specified in the value and are not impacted
  // by the value of the Prefer: outlook.timezone header if present.
  // If no timezone offset is included in the value, it is interpreted as UTC.
  // https://docs.microsoft.com/en-us/graph/api/calendar-list-calendarview?view=graph-rest-1.0&tabs=http#query-parameters
  const offset = date.getTimezoneOffset();
  const dateString = new Date(date.getTime() - offset * 60 * 1000).toISOString().slice(0, 23);
  return dateString;
}
