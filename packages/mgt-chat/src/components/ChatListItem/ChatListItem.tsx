import React, { useState, useEffect } from 'react';
import { makeStyles, mergeClasses, shorthands } from '@fluentui/react-components';
import {
  Chat,
  ConversationMember,
  AadUserConversationMember,
  NullableOption,
  ChatMessageInfo,
  TeamworkApplicationIdentity
} from '@microsoft/microsoft-graph-types';
import { Person, PersonCardInteraction } from '@microsoft/mgt-react';
import { ProviderState, Providers, error } from '@microsoft/mgt-element';
import { ChatListItemIcon } from '../ChatListItemIcon/ChatListItemIcon';
import { rewriteEmojiContent } from '../../utils/rewriteEmojiContent';
import { convert } from 'html-to-text';
import { loadChatWithPreview } from '../../statefulClient/graph.chat';
import { DefaultProfileIcon } from './DefaultProfileIcon';

interface IMgtChatListItemProps {
  chat: Chat;
  myId: string | undefined;
  isSelected: boolean;
  isRead: boolean;
}

interface EventMessageDetailWithType {
  '@odata.type': string;
}

const useStyles = makeStyles({
  // highlight selection
  isSelected: {
    backgroundColor: '#e6f7ff'
  },

  isUnSelected: {
    backgroundColor: '#ffffff'
  },

  // highlight text
  isBold: {
    fontWeight: 'bold'
  },

  isNormal: {
    fontWeight: 'normal'
  },

  chatListItem: {
    display: 'flex',
    alignItems: 'center',
    width: '100%',
    paddingRight: '10px',
    paddingLeft: '10px'
  },

  profileImage: {
    ...shorthands.flex('0 0 auto'),
    marginRight: '10px',
    objectFit: 'cover',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center'
  },

  defaultProfileImage: {
    ...shorthands.borderRadius('50%'),
    objectFit: 'cover',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center'
  },

  chatInfo: {
    flexGrow: 1,
    flexShrink: 1,
    minWidth: 0,
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
    width: 'auto'
  },

  chatMessage: {
    fontSize: '0.9em',
    color: '#666',
    textAlign: 'left',
    ...shorthands.margin('0'),
    textOverflow: 'ellipsis',
    ...shorthands.overflow('hidden'),
    whiteSpace: 'nowrap',
    width: 'auto'
  },

  chatTimestamp: {
    flexShrink: 0,
    textAlign: 'right',
    alignSelf: 'start',
    marginLeft: 'auto',
    paddingLeft: '10px',
    fontSize: '0.8em',
    color: '#999',
    whiteSpace: 'nowrap'
  },

  person: {
    '--person-avatar-size': '32px',
    '--person-alignment': 'center'
  }
});

// regex to match different tags
const imageTagRegex = /(<img[^>]+)/;
const attachmentTagRegex = /(<attachment[^>]+)/;
const systemEventTagRegex = /(<systemEventMessage[^>]+)/;

export const ChatListItem = ({ chat, myId, isSelected, isRead }: IMgtChatListItemProps) => {
  const styles = useStyles();

  // manage the internal state of the chat
  const [chatInternal, setChatInternal] = useState(chat);
  const [read, setRead] = useState<boolean>(isRead);

  // when isSelected changes to true, setRead to true
  useEffect(() => {
    if (isSelected) {
      setRead(true);
    }
  }, [isSelected]);

  // if chat changes, update the internal state to match
  useEffect(() => {
    setChatInternal(chat);
  }, [chat]);

  // enrich the chat if necessary
  useEffect(() => {
    if (chatInternal.id && (!chatInternal.chatType || !chatInternal.members)) {
      const provider = Providers.globalProvider;
      if (provider && provider.state === ProviderState.SignedIn) {
        const graph = provider.graph.forComponent('ChatListItem');
        const load = (id: string): Promise<Chat> => {
          return loadChatWithPreview(graph, id);
        };
        load(chatInternal.id).then(
          c => setChatInternal(c),
          e => error(e)
        );
      }
    }
  }, [chatInternal]);

  // shortcut if no valid user
  if (!myId) {
    return <></>;
  }

  // Copied and modified from the sample ChatItem.tsx
  // Determines the title in the case of 1:1 and self chats
  // Self Chats are not possible, however, 1:1 chats with a bot will show no other members other than self.
  const inferTitle = (c: Chat) => {
    const name = (member: ConversationMember): string => {
      return `${member?.displayName || (member as AadUserConversationMember)?.email || member?.id}`;
    };

    // build others array
    const others = (c.members || []).filter(m => (m as AadUserConversationMember).userId !== myId);
    const application = c.lastMessagePreview?.from?.application as TeamworkApplicationIdentity;

    // return the appropriate title
    if (c.topic) {
      return c.topic;
    } else if (others.length === 0 && application) {
      return application.displayName ? `${application.displayName} (bot)` : `${application.id} (bot)`;
    } else if (others.length === 0) {
      const me = c.members?.find(m => (m as AadUserConversationMember).userId === myId);
      return me ? name(me) : 'Me';
    } else if (others.length === 1) {
      return name(others[0]);
    } else if (others.length === 2) {
      return others.map(m => m.displayName?.split(' ')[0]).join(' and ');
    } else if (others.length === 3) {
      return others.map(m => m.displayName?.split(' ')[0]).join(', ');
    } else if (others.length > 3) {
      const firstThreeMembersSlice = others.slice(0, 3);
      const remainingMembersCount = others.length - 3 + 1; // +1 for the current user
      return firstThreeMembersSlice.map(m => m.displayName?.split(' ')[0]).join(', ') + ' +' + remainingMembersCount;
    } else {
      return c.id;
    }
  };

  // Derives the timestamp to display
  const extractTimestamp = (timestamp: NullableOption<string>): string => {
    if (!timestamp) return '';
    const currentDate = new Date();
    const date = new Date(timestamp);

    const [month, day, year] = [date.getMonth(), date.getDate(), date.getFullYear()];
    const [currentMonth, currentDay, currentYear] = [
      currentDate.getMonth(),
      currentDate.getDate(),
      currentDate.getFullYear()
    ];

    // if the message was sent today, return the time
    if (currentDay === day && currentMonth === month && currentYear === year) {
      return date.toLocaleTimeString(navigator.language, { hour: 'numeric', minute: '2-digit' });
    }

    // if the message was sent in a previous year, include the year
    if (currentYear !== year) {
      return date.toLocaleDateString(navigator.language, { month: 'numeric', day: 'numeric', year: 'numeric' });
    }

    // otherwise, return the month and day
    return date.toLocaleDateString(navigator.language, { month: 'numeric', day: 'numeric' });
  };

  // Chooses the correct timestamp to display
  const determineCorrectTimestamp = (c: Chat) => {
    let timestamp: NullableOption<string>;

    // lastMessageTime is the time of the last message sent in the chat
    // lastUpdatedTime is Date and time at which the chat was renamed or list of members were last changed.
    const lastMessageTime = c.lastMessagePreview?.createdDateTime
      ? new Date(c.lastMessagePreview.createdDateTime)
      : null;
    const lastUpdatedTime = c.lastUpdatedDateTime ? new Date(c.lastUpdatedDateTime) : null;

    if (lastMessageTime && lastUpdatedTime && lastMessageTime > lastUpdatedTime) {
      timestamp = String(lastMessageTime);
    } else if (lastMessageTime && lastUpdatedTime && lastUpdatedTime > lastMessageTime) {
      timestamp = String(lastUpdatedTime);
    } else if (lastMessageTime) {
      timestamp = String(lastMessageTime);
    } else if (lastUpdatedTime) {
      timestamp = String(lastUpdatedTime);
    } else {
      timestamp = null;
    }

    return timestamp;
  };

  const getDefaultProfileImage = (c: Chat) => {
    // define the JSX for FluentUI Icons + Styling
    const oneOnOneProfilePicture = <ChatListItemIcon chatType="oneOnOne" />;
    const GroupProfilePicture = <ChatListItemIcon chatType="group" />;

    const other = c.members?.find(m => (m as AadUserConversationMember).userId !== myId);
    const otherAad = other as AadUserConversationMember;

    switch (true) {
      case c.chatType === 'oneOnOne':
        return (
          <Person
            className={styles.person}
            userId={otherAad?.userId ?? undefined}
            avatarSize="small"
            showPresence={true}
            personCardInteraction={PersonCardInteraction.hover}
          >
            <DefaultProfileIcon template="no-data" />
          </Person>
        );
      case c.chatType === 'group':
        return GroupProfilePicture;
      default:
        return oneOnOneProfilePicture;
    }
  };

  const enrichPreviewMessage = (previewMessage: NullableOption<ChatMessageInfo> | undefined) => {
    let previewString = '';
    let content = previewMessage?.body?.content;

    // handle null or undefined content
    if (!content) {
      if (previewMessage?.isDeleted) {
        previewString = 'This message was deleted';
      } else if (previewMessage?.from?.user?.id === myId) {
        previewString = 'You: Sent a message';
      } else if (previewMessage?.from?.user?.displayName) {
        previewString = `${previewMessage.from.user.displayName}: Sent a message`;
      } else if (previewMessage?.from?.application?.displayName) {
        previewString = `${previewMessage.from.application.displayName}: Sent a message`;
      }
      return previewString;
    }

    // handle HTML
    if (previewMessage?.body?.contentType === 'html') {
      // handle emojis
      content = rewriteEmojiContent(content);

      // handle images
      const imageMatch = content.match(imageTagRegex);
      if (imageMatch) {
        content = 'Sent an image';
      }

      /* handle attachments
       *
       * NOTE: There is a discrepency between what the Graph API returns and what the Graph subscription
       * notification sends in the case of attachments. The Graph API returns an empty content string so
       * the preview will be 'Sent a message' whereas the Graph subscription notification returns HTML
       * content with an attachment and will display 'Sent a file'. To make this consistent, we could
       * change the below to 'Sent a message', however, the Teams client uses 'Sent a file'.
       */
      const attachmentMatch = content.match(attachmentTagRegex);
      if (attachmentMatch) {
        content = 'Sent a file';
      }

      /* handle system events
       *
       * NOTE: While Graph subscription notifications send events for both users being added and removed,
       * the Graph API only returns a message for users being added.
       */
      const systemEventMessage = content.match(systemEventTagRegex);
      if (systemEventMessage) {
        const type = (previewMessage?.eventDetail as EventMessageDetailWithType)?.['@odata.type'];
        switch (type) {
          case '#microsoft.graph.membersAddedEventMessageDetail':
            content = 'User added';
            break;
        }
      }

      // convert html to text
      content = convert(content);
    }

    // handle general chats from people and bots
    if (previewMessage?.from?.user?.id === myId) {
      previewString = 'You: ' + content;
    } else if (previewMessage?.from?.user?.displayName) {
      previewString = previewMessage?.from?.user?.displayName + ': ' + content;
    } else if (previewMessage?.from?.application?.displayName) {
      previewString = previewMessage?.from?.application?.displayName + ': ' + content;
    } else {
      previewString = content;
    }

    return previewString;
  };

  const container = mergeClasses(
    styles.chatListItem,
    isSelected ? styles.isSelected : styles.isUnSelected,
    read ? styles.isNormal : styles.isBold
  );

  return (
    <div className={container}>
      <div className={styles.profileImage}>{getDefaultProfileImage(chatInternal)}</div>
      <div className={styles.chatInfo}>
        <p className={styles.chatTitle}>{inferTitle(chatInternal)}</p>
        <p className={styles.chatMessage}>{enrichPreviewMessage(chatInternal.lastMessagePreview)}</p>
      </div>
      <div className={styles.chatTimestamp}>{extractTimestamp(determineCorrectTimestamp(chatInternal))}</div>
    </div>
  );
};
