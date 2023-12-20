import React, { useCallback } from 'react';
import { makeStyles, shorthands, Button } from '@fluentui/react-components';
import {
  Chat,
  AadUserConversationMember,
  MembersAddedEventMessageDetail,
  NullableOption,
  ChatMessageInfo
} from '@microsoft/microsoft-graph-types';
import { error } from '@microsoft/mgt-element';
import { Providers, ProviderState } from '@microsoft/mgt-element';
import { ChatListItemIcon } from '../ChatListItemIcon/ChatListItemIcon';

export interface IChatListItemInteractionProps {
  onSelected: (e: Chat) => void;
}

interface IMgtChatListItemProps {
  chat: Chat;
  myId: string | undefined;
}

const useStyles = makeStyles({
  chatListItem: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between', // Add this if you want to push the timestamp to the end
    ...shorthands.padding('10px'),
    ...shorthands.borderBottom('1px solid #ccc')
  },
  profileImage: {
    width: '60px', // Increase the width for a bigger image
    height: '60px', // Increase the height for a bigger image
    ...shorthands.borderRadius('50%'), // This will make it round
    marginRight: '10px',
    objectFit: 'cover', // This ensures the image covers the area without stretching
    position: 'relative', // Ensures the ::before pseudo-element is positioned relative to the profileImg
    display: 'flex',
    ...shorthands.overflow('hidden'), // Ensures no part of the ::before spills out
    alignItems: 'center', // This will vertically center the image
    justifyContent: 'center' // This will horizontally center the image
  },
  chatInfo: {
    alignSelf: 'left',
    alignItems: 'center',
    ...shorthands.padding('5px')
  },
  chatTitle: {
    textAlign: 'left',
    ...shorthands.margin('0'),
    fontSize: '1em',
    color: '#333',
    textOverflow: 'ellipsis',
    ...shorthands.overflow('hidden'),
    whiteSpace: 'nowrap',
    width: '300px'
  },
  chatMessage: {
    textAlign: 'left',
    ...shorthands.margin('0'),
    fontSize: '0.9em',
    color: '#666',
    textOverflow: 'ellipsis',
    ...shorthands.overflow('hidden'),
    whiteSpace: 'nowrap',
    width: '300px'
  },
  chatTimestamp: {
    textAlign: 'right',
    alignSelf: 'start',
    marginLeft: 'auto',
    paddingLeft: '10px',
    fontSize: '0.8em',
    color: '#999',
    width: '60px'
  }
});

export const ChatListItem = ({ chat, myId, onSelected }: IMgtChatListItemProps & IChatListItemInteractionProps) => {
  const styles = useStyles();

  // shortcut if no valid user
  if (!myId) {
    return <></>;
  }

  // Copied and modified from the sample ChatItem.tsx
  // Determines the title in the case of 1:1 and self chats
  const inferTitle = (chatObj: Chat) => {
    if (myId && chatObj.chatType === 'oneOnOne' && chatObj.members) {
      const other = chatObj.members.find(m => (m as AadUserConversationMember).userId !== myId);
      const me = chatObj.members.find(m => (m as AadUserConversationMember).userId === myId);
      return other
        ? `${other?.displayName || (other as AadUserConversationMember)?.email || other?.id}`
        : `${me?.displayName} (You)`;
    }
    return chatObj.topic || chatObj.chatType;
  };

  const extractTimestamp = (timestamp: NullableOption<string> | undefined): string => {
    if (timestamp === undefined || timestamp === null) return '';
    const currentDate = new Date();
    const date = new Date(timestamp);

    const [month, day, year] = [date.getMonth(), date.getDate(), date.getFullYear()];
    const [currentMonth, currentDay, currentYear] = [
      currentDate.getMonth(),
      currentDate.getDate(),
      currentDate.getFullYear()
    ];

    if (currentDay === day && currentMonth === month && currentYear === year) {
      return date.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
    }

    return date.toLocaleDateString([], { month: 'numeric', day: 'numeric' });
  };

  const getDefaultProfileImage = () => {
    // define the JSX for FluentUI Icons + Styling
    const oneOnOneProfilePicture = <ChatListItemIcon chatType="oneOnOne" />;
    const GroupProfilePicture = <ChatListItemIcon chatType="group" />;

    switch (true) {
      case chat.chatType === 'oneOnOne':
        return oneOnOneProfilePicture;
      case chat.chatType === 'group':
        return GroupProfilePicture;
      default:
        error(`Error: Unexpected chatType: ${chat.chatType}`);
        return oneOnOneProfilePicture;
    }
  };

  const removeHTMLTags = (str: string) => {
    return str.replace(/<[^>]*>/g, '');
  };

  const enrichPreviewMessage = (previewMessage: NullableOption<ChatMessageInfo> | undefined) => {
    let previewString = '';

    // this should be refactored
    // handle general chats from people and bots
    if (previewMessage?.from?.user?.id === myId) {
      previewString = 'You: ' + previewMessage?.body?.content;
    } else if (previewMessage?.from?.user?.displayName) {
      previewString = previewMessage?.from?.user?.displayName + ': ' + previewMessage?.body?.content;
    } else if (previewMessage?.from?.application?.displayName) {
      previewString = previewMessage?.from?.application?.displayName + ': ' + previewMessage?.body?.content;
    }

    // handle all events
    if (previewMessage?.eventDetail) {
      previewString = previewMessage?.body?.content as string;
    }

    return removeHTMLTags(previewString);
  };

  return (
    <Button
      className={styles.chatListItem}
      onClick={() => {
        onSelected(chat);
      }}
    >
      {getDefaultProfileImage()}
      <div className={styles.chatInfo}>
        <h3 className={styles.chatTitle}>{inferTitle(chat)}</h3>
        <p className={styles.chatMessage}>{enrichPreviewMessage(chat.lastMessagePreview)}</p>
      </div>
      <div className={styles.chatTimestamp}>{extractTimestamp(chat.lastUpdatedDateTime)}</div>
    </Button>
  );
};
