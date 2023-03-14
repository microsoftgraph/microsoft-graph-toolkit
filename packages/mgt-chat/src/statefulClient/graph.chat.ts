import { IGraph, prepScopes } from '@microsoft/mgt-element';
import { Chat, ChatMessage } from '@microsoft/microsoft-graph-types';

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
  sendChatMessage: ['chatmessage.send'],
  updateChatMessage: ['chat.readwrite'],
  deleteChatMessage: ['chat.readwrite']
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

  return graph
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
export const deleteChatMessage = async (graph: IGraph, chatId: string, messageId: string): Promise<void> =>
  graph
    .api(`/me/chats/${chatId}/messages/${messageId}/softDelete`)
    .middlewareOptions(prepScopes(...chatOperationScopes.deleteChatMessage))
    .post({});
