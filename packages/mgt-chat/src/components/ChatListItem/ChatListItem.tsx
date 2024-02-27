import React, { useState, useEffect, useRef } from 'react';
import { makeStyles, mergeClasses, shorthands } from '@fluentui/react-components';
import {
  Chat,
  ConversationMember,
  AadUserConversationMember,
  NullableOption,
  TeamworkApplicationIdentity,
  TeamsAppInstallation
} from '@microsoft/microsoft-graph-types';
import { ProviderState, Providers, error } from '@microsoft/mgt-element';
import { rewriteEmojiContentToText } from '../../utils/rewriteEmojiContent';
import { convert } from 'html-to-text';
import { loadBotsInChat } from '../../statefulClient/graph.chat';
import { DefaultProfileIcon } from './DefaultProfileIcon';
import { GraphChatThread } from '../../statefulClient/StatefulGraphChatListClient';

interface IChatListItemProps {
  chat: GraphChatThread;
  userId: string | undefined;
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
    ...shorthands.flex(0, 0, 'auto'),
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
    width: 'auto',
    minHeight: '20px' // ensures formatting is consistent with empty previewMessage
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

const IGNORE_BOTS_WITH_NAME = ['Updates'];

// regex to match different tags
const imageTagRegex = /(<img[^>]+)/;
const attachmentTagRegex = /(<attachment[^>]+)/;
const systemEventTagRegex = /(<systemEventMessage[^>]+)/;

export const ChatListItem = ({ chat, userId: myId, isSelected, isRead }: IChatListItemProps) => {
  const styles = useStyles();

  // manage the internal state of the chat
  const [bots, setBots] = useState<TeamsAppInstallation[]>();
  const isBotsLoadingOrLoaded = useRef(false);

  // if chat changes, the bots need to change too
  useEffect(() => {
    setBots(undefined);
    isBotsLoadingOrLoaded.current = false;
  }, [chat]);

  const startLoadingBotsInChat = async (chatId: string) => {
    // ensure this is only called once
    if (isBotsLoadingOrLoaded.current) {
      return;
    }

    // make sure there is a logged in graph provider
    const provider = Providers.globalProvider;
    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    // set to loading
    isBotsLoadingOrLoaded.current = true;

    // load the bots
    const graph = provider.graph.forComponent('ChatListItem');
    try {
      const botsResponse = await loadBotsInChat(graph, chatId);
      setBots(botsResponse.value);
    } catch (e) {
      error(e);
      setBots([]);
    }
  };

  // shortcut if no valid user
  if (!myId) {
    return <></>;
  }

  const getMemberName = (member: ConversationMember): string => {
    return member?.displayName || (member as AadUserConversationMember)?.email || member?.id || 'Unknown';
  };

  const getTitleFromNames = (names: string[], useFirstNamesIfAppropriate: boolean) => {
    names = names.sort();
    if (names.length === 1) {
      return names[0];
    }

    if (names.length === 2 && useFirstNamesIfAppropriate) {
      const firstNames = names.map(n => n.split(' ')[0]);
      return firstNames.join(' and ');
    }

    if (names.length === 2) {
      return names.join(' and ');
    }

    if (names.length === 3 && useFirstNamesIfAppropriate) {
      const firstNames = names.map(n => n.split(' ')[0]);
      return firstNames.join(', ');
    }

    if (names.length === 3) {
      return names.join(', ');
    }

    if (names.length > 3 && useFirstNamesIfAppropriate) {
      const firstNames = names.map(n => n.split(' ')[0]);
      const firstThree = firstNames.slice(0, 3);
      const remainingCount = names.length - 3 + 1; // +1 for the current user
      return firstThree.join(', ') + ' +' + remainingCount;
    }

    if (names.length > 3) {
      const firstThree = names.slice(0, 3);
      const remainingCount = names.length - 3 + 1; // +1 for the current user
      return firstThree.join(', ') + ' +' + remainingCount;
    }
  };

  // Copied and modified from the sample ChatItem.tsx
  // Determines the title in the case of 1:1 and self chats
  // Self Chats are not possible, however, 1:1 chats with a bot will show no other members other than self.
  const getTitleFromChat = (c: Chat) => {
    // use the topic if available
    if (c.topic) {
      return c.topic;
    }

    const others = (c.members || []).filter(m => (m as AadUserConversationMember).userId !== myId);

    // when there are no other people, and we have bot information to make the most informed decision
    if (others.length === 0 && bots) {
      const names = bots
        .filter(
          a =>
            a.teamsAppDefinition?.bot?.id &&
            a.teamsAppDefinition?.displayName &&
            IGNORE_BOTS_WITH_NAME.indexOf(a.teamsAppDefinition?.displayName) === -1
        )
        .map(a => a.teamsAppDefinition!.displayName!);
      if (names.length > 0) {
        const unique = names.filter((n, i) => names.indexOf(n) === i);
        return getTitleFromNames(unique, false);
      }
      const me = c.members?.find(m => (m as AadUserConversationMember).userId === myId);
      return me ? getMemberName(me) : 'Me';
    }

    const application = c.lastMessagePreview?.from?.application as TeamworkApplicationIdentity;

    // when there are no people, and we don't have any bot information but we do have application information
    if (others.length === 0 && application) {
      void startLoadingBotsInChat(c.id!);
      return application.displayName ? application.displayName : application.id;
    }

    // when there are no people, and we don't have any bot or application information
    if (others.length === 0) {
      void startLoadingBotsInChat(c.id!);
      return c.id;
    }

    // when there are people, use them
    if (others.length > 0) {
      const names = others.map(o => getMemberName(o));
      return getTitleFromNames(names, true);
    }

    return c.id;
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

  const getLastMessagePreviewPrefix = (c: Chat) => {
    // if the last message was sent by the current user, display 'You: '
    if (c.lastMessagePreview?.from?.user?.id === myId) {
      return 'You: ';
    }

    // if you are only chatting with one other person, don't display their name
    if (c.chatType === 'oneOnOne') {
      return '';
    }

    // if the last message was sent from a user, use their display name
    if (c.lastMessagePreview?.from?.user?.displayName) {
      return `${c.lastMessagePreview.from.user.displayName}: `;
    }

    // if the last message is from a bot and we have all the app info, use the app name
    if (c.lastMessagePreview?.from?.application?.id && bots) {
      const bot = bots.find(a => a.teamsAppDefinition?.bot?.id === c.lastMessagePreview?.from?.application?.id);
      if (bot?.teamsAppDefinition?.displayName) {
        return `${bot.teamsAppDefinition.displayName}: `;
      }
    }

    // if the last message is from a bot and we don't have all the app info, load it
    if (c.lastMessagePreview?.from?.application?.displayName) {
      void startLoadingBotsInChat(c.id!);
      return `${c.lastMessagePreview.from.application.displayName}: `;
    }

    return '';
  };

  const enrichPreviewMessage = (c: Chat) => {
    const prefix = getLastMessagePreviewPrefix(c);
    let content = c.lastMessagePreview?.body?.content;

    // handle null or undefined content
    if (!content) {
      if (c.lastMessagePreview?.isDeleted) {
        return 'This message was deleted';
      } else if (c.lastMessagePreview) {
        return `${prefix}Sent a message`;
      } else {
        return '';
      }
    }

    // handle HTML
    if (c.lastMessagePreview?.body?.contentType === 'html') {
      // handle images
      const imageMatch = content.match(imageTagRegex);
      if (imageMatch) {
        return `${prefix}Sent an image`;
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
        return `${prefix}Sent a file`;
      }

      /* handle system events
       *
       * NOTE: While Graph subscription notifications send events for both users being added and removed,
       * the Graph API only returns a message for users being added.
       */
      const systemEventMessage = content.match(systemEventTagRegex);
      if (systemEventMessage) {
        const type = (c.lastMessagePreview?.eventDetail as EventMessageDetailWithType)?.['@odata.type'];
        switch (type) {
          case '#microsoft.graph.membersAddedEventMessageDetail':
            return `${prefix}User added`;
        }
      }

      // convert html to text
      content = rewriteEmojiContentToText(content);
      content = convert(content);
    }

    // handle general chats from people and bots
    return `${prefix}${content}`;
  };

  const container = mergeClasses(
    styles.chatListItem,
    isSelected ? styles.isSelected : styles.isUnSelected,
    isRead ? styles.isNormal : styles.isBold
  );

  // short cut if the id is not defined
  if (!myId) {
    return <></>;
  }
  return (
    <div className={container}>
      <div className={styles.profileImage}>
        <DefaultProfileIcon chat={chat} userId={myId} />
      </div>
      <div className={styles.chatInfo}>
        <p className={styles.chatTitle}>{getTitleFromChat(chat)}</p>
        <p className={styles.chatMessage}>{enrichPreviewMessage(chat)}</p>
      </div>
      <div className={styles.chatTimestamp}>{extractTimestamp(determineCorrectTimestamp(chat))}</div>
    </div>
  );
};
