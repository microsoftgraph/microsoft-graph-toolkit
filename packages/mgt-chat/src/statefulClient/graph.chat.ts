/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  schemas,
  CachePhoto,
  getPhotoInvalidationTime,
  storePhotoInCache,
  blobToBase64
} from '@microsoft/mgt-components';
import { CacheService, IGraph, prepScopes } from '@microsoft/mgt-element';
import { ResponseType } from '@microsoft/microsoft-graph-client';
import { AadUserConversationMember, Chat, ChatMessage, ChatType } from '@microsoft/microsoft-graph-types';

/**
 * Generic collection response from graph
 */
type GraphCollection<T = any> = {
  '@odata.nextLink': string;
  nextLink: string;
  value: T[];
};

/**
 * Object representing a collection of chat messages
 */
export type MessageCollection = GraphCollection<ChatMessage>;

/**
 * Object mapping chat operations to the scopes required to perform them
 */
const chatOperationScopes: Record<string, string[]> = {
  loadChat: ['chat.readbasic'],
  loadChatMessages: ['chat.read'],
  loadChatImage: ['chat.read'],
  sendChatMessage: ['chatmessage.send'],
  updateChatMessage: ['chat.readwrite'],
  deleteChatMessage: ['chat.readwrite'],
  removeChatMemeber: ['chatmember.readwrite'],
  addChatMember: ['chatmember.readwrite'],
  createChat: ['chat.create']
};

/**
 * Provides an array of the distinct scopes required for all chat operations
 */
export const allChatScopes = Array.from(
  Object.values(chatOperationScopes).reduce((acc, scopes) => {
    scopes.forEach(s => acc.add(s));
    return acc;
  }, new Set<string>())
);

/**
 * Load the specified chat from graph with the members expanded
 *
 * @param graph authenticated graph client from mgt
 * @param chatId the id of the chat to load
 * @returns {Promise<Chat>}
 */
export const loadChat = async (graph: IGraph, chatId: string): Promise<Chat> =>
  (await graph
    .api(`/chats/${chatId}?$expand=members`)
    .middlewareOptions(prepScopes(...chatOperationScopes.loadChat))
    .get()) as Chat;

/**
 * Load the first page of messages from the specified chat
 * Will provide a nextLink to load more messages if there are more than the specified messageCount
 *
 * @param graph authenticated graph client from mgt
 * @param chatId the id of the chat to load messages from
 * @param messageCount how many messages to load per page
 * @returns {Promise<MessageCollection>} a collection of messages and a possible nextLink to load more messages
 */
export const loadChatThread = async (
  graph: IGraph,
  chatId: string,
  messageCount: number
): Promise<MessageCollection> => {
  const response = (await graph
    .api(`/chats/${chatId}/messages`)
    .orderby('createdDateTime DESC')
    .top(messageCount)
    .middlewareOptions(prepScopes(...chatOperationScopes.loadChatMessages))
    .get()) as MessageCollection;
  // split the nextLink on version to maintain a relative path
  response.nextLink = response['@odata.nextLink']?.split(graph.version)[1];
  return response;
};

/**
 * Loads another page of messages from the specified chat
 *
 * @param graph authenticated graph client from mgt
 * @param nextLink the nextLink to load more messages from
 * @returns {Promise<MessageCollection>} a collection of messages and a possilbe nextLink to load more messages
 */
export const loadMoreChatMessages = async (graph: IGraph, nextLink: string): Promise<MessageCollection> => {
  // not defining scopes here because this uses a nextLink and the initial request ensures the required scopes are present
  const response = (await graph.api(nextLink).get()) as MessageCollection;
  // split the nextLink on version to maintain a relative path
  response.nextLink = response['@odata.nextLink']?.split(graph.version)[1];
  return response;
};

/**
 * Creates a new chat message via HTTP POST
 *
 * @param graph authenticated graph client from mgt
 * @param chatId id of the chat to send the message to
 * @param content content of the message to send
 * @returns {Promise<ChatMessage>} the newly created message
 */
export const sendChatMessage = async (graph: IGraph, chatId: string, content: string): Promise<ChatMessage> => {
  // TODO: remove this code that lets me simulate a failure during debugging
  // let fail = false;

  // debugger;
  // if (fail) throw new Error('fail');

  return (await graph
    .api(`/chats/${chatId}/messages`)
    .middlewareOptions(prepScopes(...chatOperationScopes.sendChatMessage))
    .post({ body: { content } })) as ChatMessage;
};

/**
 * Updates a chat message via HTTP PATCH
 *
 * @param graph authenticated graph client from mgt
 * @param chatId the id of the chat containing the message to update
 * @param messageId the id of the message to update
 * @param content the new content of the message
 * @returns {Promise<void>}
 */
export const updateChatMessage = async (
  graph: IGraph,
  chatId: string,
  messageId: string,
  content: string
): Promise<void> => {
  // TODO: remove this code that lets me simulate a failure during debugging
  // let fail = false;

  // debugger;
  // if (fail) throw new Error('fail');

  await graph
    .api(`/chats/${chatId}/messages/${messageId}`)
    .middlewareOptions(prepScopes(...chatOperationScopes.updateChatMessage))
    .patch({ body: { content } });
};

/**
 * Deletes a chat message via HTTP POST to the softDelete action
 *
 * @param graph authenticated graph client from mgt
 * @param chatId the id of the chat containing the message to delete
 * @param messageId the id of the message to delete
 * @returns {Promise<void>}
 */
export const deleteChatMessage = async (graph: IGraph, chatId: string, messageId: string): Promise<void> => {
  await graph
    .api(`/me/chats/${chatId}/messages/${messageId}/softDelete`)
    .middlewareOptions(prepScopes(...chatOperationScopes.deleteChatMessage))
    .post({});
};

export const removeChatMember = async (graph: IGraph, chatId: string, membershipId: string): Promise<void> => {
  await graph
    .api(`/chats/${chatId}/members/${membershipId}`)
    .middlewareOptions(prepScopes(...chatOperationScopes.deleteChatMessage))
    .delete();
};

export const addChatMembers = async (
  graph: IGraph,
  chatId: string,
  userIds: string[],
  visibleHistoryStartDateTime?: Date
): Promise<void> => {
  const body = {
    requests: userIds.map((userId: string, index: number) => {
      const addRequestBody: AadUserConversationMember & { '@odata.type': string; 'user@odata.bind': string } = {
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${userId}`,
        roles: ['owner']
      };
      if (visibleHistoryStartDateTime) {
        addRequestBody.visibleHistoryStartDateTime = visibleHistoryStartDateTime.toISOString();
      }
      return {
        id: index,
        method: 'POST',
        url: `/chats/${chatId}/members`,
        headers: {
          'Content-Type': 'application/json'
        },
        body: addRequestBody
      };
    })
  };

  await graph
    .api('$batch')
    .middlewareOptions(prepScopes(...chatOperationScopes.addChatMember))
    .post(body);
};
/**
 * Whether or not the cache is enabled
 */
export const isPhotoCacheEnabled = (): boolean => CacheService.config.photos.isEnabled && CacheService.config.isEnabled;

export const loadChatImage = async (graph: IGraph, url: string): Promise<string | null> => {
  let cachedPhoto: CachePhoto;

  // attempt to get user and photo from cache if enabled
  if (isPhotoCacheEnabled()) {
    const cache = CacheService.getCache<CachePhoto>(schemas.photos, schemas.photos.stores.teams);
    cachedPhoto = await cache.getValue(url);
    if (
      cachedPhoto?.timeCached &&
      cachedPhoto.photo &&
      getPhotoInvalidationTime() > Date.now() - cachedPhoto.timeCached
    ) {
      return cachedPhoto.photo;
    }
  }
  const response = (await graph
    .api(url)
    .responseType(ResponseType.RAW)
    .middlewareOptions(prepScopes(...chatOperationScopes.loadChatImage))
    .get()) as Response & { '@odata.mediaEtag'?: string };

  if (response.status === 404) {
    // 404 means the resource does not have a photo
    // we still want to cache that state
    // so we return an object that can be cached
    cachedPhoto = { eTag: undefined, photo: undefined };
  } else if (!response.ok) {
    return null;
  }

  const eTag = response['@odata.mediaEtag'];
  const blob = await blobToBase64(await response.blob());

  cachedPhoto = { eTag, photo: blob };
  await storePhotoInCache(url, schemas.photos.stores.teams, cachedPhoto);
  return blob;
};

/**
 * Creates a new chat thread via HTTP POST
 *
 * @param graph authenticated graph client from mgt
 * @param memberIds array of the user ids of users to be members of the chat. Must be at least 2.
 * @param isGroupChat if true a group chat will be created, otherwise a 1:1 chat will be created. Must be true if there are more than 2 members.
 * @param chatMessage if provided a message with this content will be sent to the chat after creation
 * @param chatName if provided and there are more than 2 members the topic of the chat will be set to this value
 * @returns {Promise<Chat>} the newly created chat
 */
export const createChatThread = async (
  graph: IGraph,
  memberIds: string[],
  isGroupChat: boolean,
  chatMessage: string | undefined,
  chatName: string | undefined
): Promise<Chat> => {
  const body: Chat = {
    chatType: isGroupChat || memberIds.length > 2 ? 'group' : 'oneOnOne',
    members: memberIds.map(id => {
      return {
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles: ['owner'],
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${id}')`
      };
    })
  };
  if (isGroupChat && chatName) body.topic = chatName;

  const chat = (await graph
    .api('/chats')
    .middlewareOptions(prepScopes(...chatOperationScopes.loadChatImage))
    .post(body)) as Chat;
  if (!chat?.id) throw new Error('Chat id not returned from create chat thread');
  if (chatMessage) {
    await sendChatMessage(graph, chat.id, chatMessage);
  }

  return chat;
};

/**
 * Updates the topic of a chat via HTTP PATCH
 *
 * @param graph {IGraph} - authenticated graph client from mgt
 * @param chatId {string} - id of the chat to update
 * @param topic {string | null} - new value for the chat topic, null will remove the currently set topic
 */
export const updateChatTopic = async (graph: IGraph, chatId: string, topic: string | null): Promise<void> => {
  await graph
    .api(`/chats/${chatId}`)
    .middlewareOptions(prepScopes(...chatOperationScopes.updateChatMessage))
    .patch({ topic });
};
