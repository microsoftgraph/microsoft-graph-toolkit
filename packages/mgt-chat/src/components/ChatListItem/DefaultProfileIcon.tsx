import React, { useCallback } from 'react';
import { makeStyles } from '@fluentui/react-components';
import { ChatListItemIcon } from '../ChatListItemIcon/ChatListItemIcon';
import { Chat, AadUserConversationMember } from '@microsoft/microsoft-graph-types';
import { Person } from '@microsoft/mgt-react';
import { MgtTemplateProps } from '@microsoft/mgt-react';

interface IDefaultProfileIconProps {
  chat: Chat;
  userId: string;
}

const useStyles = makeStyles({
  person: {
    '--person-avatar-size': '32px',
    '--person-alignment': 'center'
  }
});

export const DefaultProfileIcon = ({ chat, userId }: IDefaultProfileIconProps & MgtTemplateProps) => {
  const styles = useStyles();

  const otherAad = useCallback(() => {
    const other = chat.members?.find(m => (m as AadUserConversationMember).userId !== userId);
    return other as AadUserConversationMember;
  }, [chat, userId]);

  return (
    <>
      {chat.chatType === 'oneOnOne' && (
        <Person
          className={styles.person}
          userId={otherAad()?.userId ?? undefined}
          avatarSize="small"
          showPresence={true}
          personCardInteraction="hover"
        >
          <ChatListItemIcon chatType="oneOnOne" template="no-data" />
        </Person>
      )}
      {chat.chatType === 'group' && <ChatListItemIcon chatType="group" />}
      {chat.chatType === 'meeting' && <ChatListItemIcon chatType="meeting" />}
    </>
  );
};
