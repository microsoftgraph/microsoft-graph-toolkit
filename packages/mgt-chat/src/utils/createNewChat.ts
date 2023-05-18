import { Chat } from '@microsoft/microsoft-graph-types';
import { createChatThread } from '../statefulClient/graph.chat';
import { graph } from './graph';
/**
 * Creates a new chat thread via HTTP POST
 *
 * @param memberIds array of the user ids of users to be members of the chat. Must be at least 2.
 * @param isGroupChat if true a group chat will be created, otherwise a 1:1 chat will be created. Must be true if there are more than 2 members.
 * @param chatMessage if provided a message with this content will be sent to the chat after creation
 * @param chatName if provided and there are more than 2 members the topic of the chat will be set to this value
 * @returns {Promise<Chat>} the newly created chat
 */
export const createNewChat = (
  memberIds: string[],
  isGroupChat: boolean,
  chatMessage: string | undefined,
  chatName: string | undefined
): Promise<Chat> => {
  return createChatThread(graph('mgt-chat'), memberIds, isGroupChat, chatMessage, chatName);
};
