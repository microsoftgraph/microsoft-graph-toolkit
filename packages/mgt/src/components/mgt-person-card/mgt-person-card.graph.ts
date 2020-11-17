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
  'businessPhones,companyName,department,displayName,givenName,jobTitle,mail,mobilePhone,officeLocation,preferredLanguage,surname,userPrincipalName,id';

// tslint:disable-next-line:completed-docs
const batchKeys = {
  directReports: 'directReports',
  files: 'files',
  messages: 'messages',
  people: 'people',
  person: 'person'
};

/**
 * Get initial data to populate the person card
 *
 * @param {string} userId
 * @returns {(Promise<IDynamicPerson>)}
 * @memberof Graph
 */
export async function getPersonCardGraphData(
  graph: IGraph,
  personDetails: IDynamicPerson,
  isMe: boolean,
  config: MgtPersonCardConfig
): Promise<MgtPersonCardState> {
  const userId = personDetails.id;
  const email = getEmailFromGraphEntity(personDetails);

  const batch = graph.createBatch();

  if (config.sections.organization) {
    buildOrgStructureRequest(batch, userId);
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

  if (config.sections.profile) {
    try {
      const profile = await getProfile(graph, userId);
      if (profile) {
        data.profile = profile;
      }
    } catch {
      // nop
    }
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

  batch.get(batchKeys.files, request); // TODO , ['sites.read.all']);
}

async function getProfile(graph: IGraph, userId: string): Promise<Profile> {
  const profile = await graph
    .api(`/users/${userId}/profile`)
    .version('beta')
    .get();
  return profile;
}
