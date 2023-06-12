/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { BatchResponse, IBatch, IGraph, prepScopes } from '@microsoft/mgt-element';
import { Chat, ChatMessage } from '@microsoft/microsoft-graph-types';
import { Profile } from '@microsoft/microsoft-graph-types-beta';

import { getEmailFromGraphEntity } from '../../graph/graph.people';
import { IDynamicPerson } from '../../graph/types';
import { MgtPersonCardConfig, MgtPersonCardState } from './mgt-person-card.types';

const userProperties =
  'businessPhones,companyName,department,displayName,givenName,jobTitle,mail,mobilePhone,officeLocation,preferredLanguage,surname,userPrincipalName,id,accountEnabled';

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
export const getPersonCardGraphData = async (
  graph: IGraph,
  personDetails: IDynamicPerson,
  isMe: boolean,
  config: MgtPersonCardConfig
): Promise<MgtPersonCardState> => {
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
      // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-member-access
      data[key] = value.content?.value || value.content;
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
};

const buildOrgStructureRequest = (batch: IBatch, userId: string) => {
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
};

const buildWorksWithRequest = (batch: IBatch, userId: string) => {
  batch.get(batchKeys.people, `users/${userId}/people?$filter=personType/class eq 'Person'`, ['People.Read.All']);
};

const buildMessagesWithUserRequest = (batch: IBatch, emailAddress: string) => {
  batch.get(batchKeys.messages, `me/messages?$search="from:${emailAddress}"`, ['Mail.ReadBasic']);
};

const buildFilesRequest = (batch: IBatch, emailAddress?: string) => {
  let request: string;

  if (emailAddress) {
    request = `me/insights/shared?$filter=lastshared/sharedby/address eq '${emailAddress}'`;
  } else {
    request = 'me/insights/used';
  }

  batch.get(batchKeys.files, request, ['Sites.Read.All']);
};

/**
 * Get the profile for a user
 *
 * @param {IGraph} graph
 * @param {string} userId
 * @return {*}  {Promise<Profile>}
 */
const getProfile = async (graph: IGraph, userId: string): Promise<Profile> =>
  (await graph.api(`/users/${userId}/profile`).version('beta').get()) as Profile;

/**
 * Initiate a chat to a user
 *
 * @export
 * @param {IGraph} graph
 * @param {{ chatType: string; members: [{"@odata.type": string,"roles": ["owner"],"user@odata.bind": string},{"@odata.type": string,"roles": ["owner"],"user@odata.bind": string}]  }} chatData
 * @return {*}  {Promise<Chat>}
 */
export const createChat = async (graph: IGraph, person: string, user: string): Promise<Chat> => {
  const chatData = {
    chatType: 'oneonOne',
    members: [
      {
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles: ['owner'],
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${user}')`
      },
      {
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles: ['owner'],
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${person}')`
      }
    ]
  };
  return (await graph
    .api('/chats')
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Chat.Create', 'Chat.ReadWrite'))
    .post(chatData)) as Chat;
};

/**
 * Send a chat message to a user
 *
 * @export
 * @param {IGraph} graph
 * @param {{ body: {"content": string}  }} messageData
 * @return {*}  {Promise<ChatMessage>}
 */
export const sendMessage = async (
  graph: IGraph,
  chatId: string,
  messageData: Pick<ChatMessage, 'body'>
): Promise<ChatMessage> =>
  (await graph
    .api(`/chats/${chatId}/messages`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Chat.ReadWrite', 'ChatMessage.Send'))
    .post(messageData)) as ChatMessage;
