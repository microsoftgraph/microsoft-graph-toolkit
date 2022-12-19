/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { BatchResponse, IBatch, IGraph } from '@microsoft/mgt-element';
import { Profile } from '@microsoft/microsoft-graph-types-beta';

import { getEmailFromGraphEntity } from '../../graph/graph.people';
import { IDynamicPerson } from '../../graph/types';
import { MgtPersonCardConfig, MgtPersonCardState } from './mgt-person-card.types';

// tslint:disable-next-line:completed-docs
const userProperties =
  'businessPhones,companyName,department,displayName,givenName,jobTitle,mail,mobilePhone,officeLocation,preferredLanguage,surname,userPrincipalName,id,accountEnabled';

// tslint:disable-next-line:completed-docs
const batchKeys = {
  directReports: 'directReports',
  files: 'files',
  messages: 'messages',
  people: 'people',
  person: 'person'
};

/**
 * Get data to populate the person card
 *
 * @export
 * @param {IGraph} graph
 * @param {IDynamicPerson} personDetails
 * @param {boolean} isMe
 * @param {MgtPersonCardConfig} config
 * @return {*}  {Promise<MgtPersonCardState>}
 */
export async function getPersonCardGraphData(
  graph: IGraph,
  personDetails: IDynamicPerson,
  isMe: boolean,
  config: MgtPersonCardConfig
): Promise<MgtPersonCardState> {
  const userId = personDetails.id;
  const email = getEmailFromGraphEntity(personDetails);

  const isContactOrGroup =
    'classification' in personDetails ||
    ('personType' in personDetails &&
      (personDetails.personType.subclass === 'PersonalContact' || personDetails.personType.class === 'Group'));

  const batch = graph.createBatch();

  if (!isContactOrGroup) {
    if (config.sections.organization) {
      buildOrgStructureRequest(batch, userId);

      if (typeof config.sections.organization !== 'boolean' && config.sections.organization.showWorksWith) {
        buildWorksWithRequest(batch, userId);
      }
    }
  }

  if (config.sections.mailMessages && email) {
    buildMessagesWithUserRequest(batch, email);
  }

  if (config.sections.files) {
    buildFilesRequest(batch, isMe ? null : email);
  }

  let response: Map<string, BatchResponse>;
  const data: MgtPersonCardState = {}; // TODO
  try {
    response = await batch.executeAll();
  } catch {
    // nop
  }

  if (response) {
    for (const [key, value] of response) {
      data[key] = value.content.value || value.content;
    }
  }

  if (!isContactOrGroup && config.sections.profile) {
    try {
      const profile = await getProfile(graph, userId);
      if (profile) {
        data.profile = profile;
      }
    } catch {
      // nop
    }
  }

  // filter out disabled users from direct reports.
  if (data.directReports && data.directReports.length > 0) {
    data.directReports = data.directReports.filter(report => report.accountEnabled);
  }

  return data;
}

// tslint:disable-next-line:completed-docs
function buildOrgStructureRequest(batch: IBatch, userId: string) {
  const expandManagers = `manager($levels=max;$select=${userProperties})`;

  batch.get(
    batchKeys.person,
    `users/${userId}?$expand=${expandManagers}&$select=${userProperties}&$count=true`,
    ['user.read.all'],
    {
      ConsistencyLevel: 'eventual'
    }
  );

  batch.get(batchKeys.directReports, `users/${userId}/directReports?$select=${userProperties}`);
}

// tslint:disable-next-line:completed-docs
function buildWorksWithRequest(batch: IBatch, userId: string) {
  batch.get(batchKeys.people, `users/${userId}/people?$filter=personType/class eq 'Person'`, ['People.Read.All']);
}

// tslint:disable-next-line:completed-docs
function buildMessagesWithUserRequest(batch: IBatch, emailAddress: string) {
  batch.get(batchKeys.messages, `me/messages?$search="from:${emailAddress}"`, ['Mail.ReadBasic']);
}

// tslint:disable-next-line:completed-docs
function buildFilesRequest(batch: IBatch, emailAddress?: string) {
  let request: string;

  if (emailAddress) {
    request = `me/insights/shared?$filter=lastshared/sharedby/address eq '${emailAddress}'`;
  } else {
    request = 'me/insights/used';
  }

  batch.get(batchKeys.files, request, ['Sites.Read.All']);
}

/**
 * Get the profile for a user
 *
 * @param {IGraph} graph
 * @param {string} userId
 * @return {*}  {Promise<Profile>}
 */
async function getProfile(graph: IGraph, userId: string): Promise<Profile> {
  const profile = await graph.api(`/users/${userId}/profile`).version('beta').get();
  return profile;
}
