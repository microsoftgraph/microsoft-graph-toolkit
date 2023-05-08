import { AadUserConversationMember as GraphChatParticipant } from '@microsoft/microsoft-graph-types';

import { ChatParticipant as AcsChatParticipant } from '@azure/communication-chat';

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
