import { IGraph } from '@microsoft/mgt-element';
import { Chat, ChatMessage } from '@microsoft/microsoft-graph-types';

export type MessageCollection = {
  '@odata.nextLink': string;
  nextLink: string;
  value: ChatMessage[];
};

export const loadChat = async (graph: IGraph, chatId: string): Promise<Chat> =>
  (await graph.api(`/chats/${chatId}?$expand=members`).get()) as Chat;

export const loadChatThread = async (
  graph: IGraph,
  chatId: string,
  messageCount: number
): Promise<MessageCollection> => {
  const response = (await graph.api(`/chats/${chatId}/messages`).top(messageCount).get()) as MessageCollection;
  // split the nextLink on version to maintain a relative path
  response.nextLink = response['@odata.nextLink']?.split(graph.version)[1];
  return response;
};

export const loadMoreChatMessages = async (graph: IGraph, nextLink: string): Promise<MessageCollection> => {
  const response = (await graph.api(nextLink).get()) as MessageCollection;
  // split the nextLink on version to maintain a relative path
  response.nextLink = response['@odata.nextLink']?.split(graph.version)[1];
  return response;
};
