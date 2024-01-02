import React, { memo } from 'react';
import { Chat } from '@microsoft/microsoft-graph-types';
import { makeStyles, mergeClasses } from '@fluentui/react-components';
import { useCommonHeaderStyles } from './useCommonHeaderStyles';
import { OneToOneChatHeader } from './OneToOneChatHeader';
import { GroupChatHeader } from './GroupChatHeader';

export interface ChatHeaderProps {
  chat?: Chat;
  currentUserId?: string;
  onRenameChat: (newName: string | null) => Promise<void>;
}

const useTitleStyles = makeStyles({
  chatHeader: {
    display: 'flex',
    alignItems: 'center',
    columnGap: '0',
    lineHeight: '32px',
    fontSize: '18px',
    fontWeight: 700,
    marginBlockStart: '11px',
    marginBlockEnd: '7px',
    zIndex: 3
  }
});

/**
 * Shows the "name" of the chat.
 * For group chats where the topic is set it will show the topic.
 * For group chats with no topic it will show the first names of other participants comma separated.
 * For 1:1 chats it will show the first name of the other participant.
 */
const ChatTitle = ({ chat, currentUserId, onRenameChat }: ChatHeaderProps) => {
  const styles = useTitleStyles();
  const commonStyles = useCommonHeaderStyles();
  return (
    <div className={mergeClasses(styles.chatHeader, commonStyles.row)}>
      {chat?.chatType !== 'oneOnOne' ? (
        <GroupChatHeader chat={chat} currentUserId={currentUserId} onRenameChat={onRenameChat} />
      ) : (
        <OneToOneChatHeader chat={chat} currentUserId={currentUserId} onRenameChat={onRenameChat} />
      )}
    </div>
  );
};

export default memo(ChatTitle);
