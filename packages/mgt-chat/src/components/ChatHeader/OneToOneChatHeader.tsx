import React from 'react';
import { AadUserConversationMember, Chat } from '@microsoft/microsoft-graph-types';
import { Person, PersonCardInteraction, ViewType } from '@microsoft/mgt-react';
import { ChatHeaderProps } from './ChatTitle';
import { makeStyles } from '@fluentui/react-components';

const getOtherParticipantUserId = (chat?: Chat, currentUserId = '') =>
  (chat?.members as AadUserConversationMember[])?.find(m => m.userId !== currentUserId)?.userId;

const useStyles = makeStyles({
  person: {
    '--person-avatar-size': '32px',
    '--person-alignment': 'center'
  }
});
export const OneToOneChatHeader = ({ chat, currentUserId }: ChatHeaderProps) => {
  const styles = useStyles();
  const id = getOtherParticipantUserId(chat, currentUserId);
  return id ? (
    <Person
      className={styles.person}
      userId={id}
      view={ViewType.oneline}
      avatarSize="small"
      personCardInteraction={PersonCardInteraction.hover}
      showPresence={true}
    />
  ) : null;
};
