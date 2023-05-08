import React, { memo } from 'react';
import { AadUserConversationMember, Chat } from '@microsoft/microsoft-graph-types';
import { styles } from './chat-header.styles';
import { IconButton, IIconProps } from '@fluentui/react';
import { Person, PersonCardInteraction, ViewType } from '@microsoft/mgt-react';
import { buttonIconStyle, buttonIconStyles, iconOnlyButtonStyle } from '../styles/common.styles';

interface ChatHeaderProps {
  chat?: Chat;
  currentUserId?: string;
}

const reduceToFirstNamesList = (participants: AadUserConversationMember[] = [], userId = '') => {
  return participants
    .filter(p => p.userId !== userId)
    .map(p => {
      if (p.displayName?.includes(' ')) {
        return p.displayName.split(' ')[0];
      }
      return p.displayName || p.email || p.id;
    })
    .join(', ');
};

const GroupChatHeader = ({ chat, currentUserId }: ChatHeaderProps) => {
  const chatTitle = chat?.topic
    ? chat.topic
    : reduceToFirstNamesList(chat?.members as AadUserConversationMember[], currentUserId);

  const editIcon: IIconProps = {
    iconName: 'edit-svg',
    className: (buttonIconStyles.button, iconOnlyButtonStyle),
    styles: { root: buttonIconStyle }
  };
  return (
    <>
      {chatTitle}
      <IconButton
        iconProps={editIcon}
        className={(buttonIconStyles.button, iconOnlyButtonStyle)}
        ariaLabel="Name group chat"
      />
    </>
  );
};

const OneToOneChatHeader = ({ chat, currentUserId }: ChatHeaderProps) => {
  const id = getOtherParticipantUserId(chat, currentUserId);
  return id ? (
    <Person
      userId={id}
      view={ViewType.oneline}
      avatarSize="small"
      personCardInteraction={PersonCardInteraction.hover}
      showPresence={true}
    />
  ) : null;
};

const getOtherParticipantUserId = (chat?: Chat, currentUserId = '') =>
  (chat?.members as AadUserConversationMember[])?.find(m => m.userId !== currentUserId)?.userId;

/**
 * Shows the "name" of the chat.
 * For group chats where the topic is set it will show the topic.
 * For group chats with no topic it will show the first names of other participants comma separated.
 * For 1:1 chats it will show the first name of the other participant.
 */
const ChatHeader = memo(({ chat, currentUserId }: ChatHeaderProps) => {
  return (
    <div className={styles.chatHeader}>
      {chat?.chatType === 'group' ? (
        <GroupChatHeader chat={chat} currentUserId={currentUserId} />
      ) : (
        <OneToOneChatHeader chat={chat} currentUserId={currentUserId} />
      )}
    </div>
  );
});

export default ChatHeader;
