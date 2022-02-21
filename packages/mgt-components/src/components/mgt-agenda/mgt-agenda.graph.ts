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
  console.log({ sdt, edt });

  let uri: string;

  if (groupId) {
    uri = `groups/${groupId}/calendar`;
  } else {
    uri = 'me';
  }

  if (preferredTimezone) {
    sdt = sdt.slice(0, -14);
    edt = edt.slice(0, -14);
  }

  uri += `/calendarview?${sdt}&${edt}`;

  let request = graph.api(uri).middlewareOptions(prepScopes(scopes)).orderby('start/dateTime');

  if (preferredTimezone) {
    request = request.header('Prefer', `outlook.timezone="${preferredTimezone}"`);
  }

  return GraphPageIterator.create<MicrosoftGraph.Event>(graph, request);
}

function dateToLocalISO(date: Date): string {
  const off = date.getTimezoneOffset();
  const absoff = Math.abs(off);
  const dateString = new Date(date.getTime() - off * 60 * 1000).toISOString().slice(0, 23);
  const offSetSign = off > 0 ? '-' : '+';
  const secs = Math.floor(absoff / 60)
    .toFixed(0)
    .padStart(2, '0');
  const ms = (absoff % 60).toString().padStart(2, '0');
  const tail = `${offSetSign}${secs}${ms}`;
  console.log({ off, secs, ms, off2: encodeURIComponent(offSetSign) });
  const tzOffSetDateString = `${dateString}${tail}`;
  return encodeURIComponent(tzOffSetDateString);
}
