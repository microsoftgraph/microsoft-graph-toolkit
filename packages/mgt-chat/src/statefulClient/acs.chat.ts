import {
  ChatMessage as GraphChatMessage,
  AadUserConversationMember as GraphChatParticipant,
  MembersAddedEventMessageDetail
} from '@microsoft/microsoft-graph-types';

import { ChatParticipant as AcsChatParticipant } from '@azure/communication-chat';
import { ChatMessage as ACSChatMessage, CommunicationParticipant, SystemMessage } from '@azure/communication-react';

export const graphParticipantToAcsParticipant = (graphParticipant: GraphChatParticipant): AcsChatParticipant => {
  if (!graphParticipant.userId) {
    throw new Error('Cannot convert graph participant to ACS participant. No ID found on graph participant');
  }

  return {
    id: {
      id: graphParticipant.userId
    },
    displayName: graphParticipant.displayName ?? undefined
  };
};

export const graphChatMessageToAcsChatMessage = (
  graphMessage: GraphChatMessage,
  currentUser: string
): ACSChatMessage => {
  if (!graphMessage.id) {
    throw new Error('Cannot convert graph message to ACS message. No ID found on graph message');
  }
  const senderId = graphMessage.from?.user?.id || undefined;
  return {
    messageId: graphMessage.id,
    contentType: graphMessage.body?.contentType ?? 'text',
    messageType: 'chat',
    content: graphMessage.body?.content ?? 'undefined',
    senderDisplayName: graphMessage.from?.user?.displayName ?? undefined,
    createdOn: new Date(graphMessage.createdDateTime ?? Date.now()),
    editedOn: graphMessage.lastEditedDateTime ? new Date(graphMessage.lastEditedDateTime) : undefined,
    senderId,
    mine: senderId === currentUser,
    status: 'seen',
    attached: 'top' // this ensures that the user's avatar is shown on messages
  };
};
