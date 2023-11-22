import {
  FluentThemeProvider,
  MessageThread,
  SendBox,
  MessageThreadStyles,
  MentionLookupOptions,
  Mention
} from '@azure/communication-react';
import { FluentTheme } from '@fluentui/react';
import { FluentProvider, makeStyles, shorthands, webLightTheme } from '@fluentui/react-components';
import { Person, Spinner } from '@microsoft/mgt-react';
import React, { useEffect, useState } from 'react';
import { StatefulGraphChatClient } from '../../statefulClient/StatefulGraphChatClient';
import { useGraphChatClient } from '../../statefulClient/useGraphChatClient';
import { onRenderMessage } from '../../utils/chat';
import { renderMGTMention } from '../../utils/mentions';
import { registerAppIcons } from '../styles/registerIcons';
import { ChatHeader } from '../ChatHeader/ChatHeader';
import { findUsers } from '@microsoft/mgt-components';
import { graph } from '../../utils/graph';
import { User } from '@microsoft/microsoft-graph-types';

registerAppIcons();

interface IMgtChatProps {
  chatId: string;
}

const useStyles = makeStyles({
  chat: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    ...shorthands.overflow('auto'),
    paddingBlockEnd: '12px',
    backgroundColor: 'var(--Neutral-Background-2-Rest, #FAFAFA)'
  },
  chatMessages: {
    height: 'auto',
    ...shorthands.paddingInline('20px'),
    ...shorthands.overflow('auto'),
    '& img': {
      maxWidth: '100%',
      height: 'auto'
    },

    '& ul': {
      ...shorthands.padding('unset')
    },
    '& .ui-chat__item__message': {
      marginLeft: 'unset',
      '& ul': {
        ...shorthands.paddingInline('40px', '0px')
      }
    },
    '& .fui-Chat': {
      maxWidth: 'unset'
    }
  },
  chatInput: {
    ...shorthands.paddingInline('16px'),
    ...shorthands.overflow('unset')
  },
  fullHeight: {
    height: '100%'
  },
  spinner: {
    justifyContent: 'center',
    display: 'flex',
    alignItems: 'center',
    height: '100%'
  },
  unsupportedContent: {
    color: 'red'
  }
});

/**
 * Styling for the MessageThread and its components.
 */
const messageThreadStyles: MessageThreadStyles = {
  chatContainer: {
    '& .ui-box': {
      zIndex: 'unset',
      '& div[data-ui-status]': {
        display: 'inline-flex',
        justifyContent: 'center'
      }
    }
  },
  chatMessageContainer: {
    '& p>.mgt-person-mention,msft-mention': {
      display: 'inline-block',
      ...shorthands.marginInline('0px')
    },
    '& .otherMention': {
      color: 'var(--accent-base-color)',
      ...shorthands.margin('0px')
    }
  },
  myChatMessageContainer: {
    '& p>.mgt-person-mention,msft-mention': {
      display: 'inline-block',
      ...shorthands.marginInline('0px')
    },
    '& .otherMention': {
      color: 'var(--accent-base-color)',
      ...shorthands.margin('0px')
    }
  }
};
const onSuggestionSelectedCb = (suggested: Mention) => {
  console.log('suggestion selected ', suggested);
  return;
};
const mentionLookupOptions: MentionLookupOptions = {
  onQueryUpdated: (query: string): Promise<Mention[]> => {
    const getUsers = async (): Promise<Mention[]> => {
      const mentions: Mention[] = [];
      const users: User[] = await findUsers(graph('mgt-chat'), query);
      if (users) {
        users.forEach(user =>
          mentions.push({ id: user?.id ?? '', displayText: user?.displayName ?? user?.givenName ?? '' })
        );
      }
      return Promise.resolve(mentions);
    };
    // NOTE: filter only people in the chat?
    return getUsers();
  },
  onRenderSuggestionItem: (suggestion: Mention, onSuggestionSelected = onSuggestionSelectedCb): JSX.Element => {
    console.log('suggestion ', suggestion);
    console.log('selected ', onSuggestionSelected);
    // NOTE: how do I override the onSuggestionSelected callback
    return <Person userId={suggestion?.id} view="oneline" onClick={() => onSuggestionSelected}></Person>;
  }
};

export const Chat = ({ chatId }: IMgtChatProps) => {
  const styles = useStyles();
  const chatClient: StatefulGraphChatClient = useGraphChatClient(chatId);
  const [chatState, setChatState] = useState(chatClient.getState());
  useEffect(() => {
    chatClient.onStateChange(setChatState);
    return () => {
      chatClient.offStateChange(setChatState);
    };
  }, [chatClient]);

  const isLoading = ['creating server connections', 'subscribing to notifications', 'loading messages'].includes(
    chatState.status
  );

  const disabled = !chatId || !!chatState.activeErrorMessages.length;
  const placeholderText = disabled ? 'You cannot send a message' : 'Type a message...';

  return (
    <FluentThemeProvider fluentTheme={FluentTheme}>
      <FluentProvider id="fluentui" theme={webLightTheme} className={styles.fullHeight}>
        <div className={styles.chat}>
          <ChatHeader chatState={chatState} />
          {chatState.userId && chatId && chatState.messages.length > 0 ? (
            <>
              <div className={styles.chatMessages}>
                <MessageThread
                  userId={chatState.userId}
                  messages={chatState.messages}
                  showMessageDate={true}
                  disableEditing={chatState.disableEditing}
                  numberOfChatMessagesToReload={chatState.numberOfChatMessagesToReload}
                  onLoadPreviousChatMessages={chatState.onLoadPreviousChatMessages}
                  // TODO: Messages date rendering is behind beta flag, find out how to enable it
                  // onDisplayDateTimeString={(date: Date) => date.toISOString()}

                  // current behavior for re-send is a delete call with the clientMessageId and the a new send call
                  onDeleteMessage={chatState.onDeleteMessage}
                  onSendMessage={chatState.onSendMessage}
                  onUpdateMessage={chatState.onUpdateMessage}
                  // render props
                  onRenderAvatar={(userId?: string) => {
                    return (
                      <Person userId={userId} avatarSize="small" personCardInteraction="hover" showPresence={true} />
                    );
                  }}
                  styles={messageThreadStyles}
                  mentionOptions={{
                    displayOptions: {
                      onRenderMention: renderMGTMention(chatState)
                    }
                  }}
                  onRenderMessage={onRenderMessage}
                />
              </div>
              <div className={styles.chatInput}>
                <SendBox
                  mentionLookupOptions={mentionLookupOptions}
                  onSendMessage={chatState.onSendMessage}
                  strings={{ placeholderText }}
                />
              </div>
            </>
          ) : (
            <>
              {isLoading && (
                <div className={styles.spinner}>
                  <Spinner /> <br />
                  {chatState.status}
                </div>
              )}
              {chatState.status === 'no messages' && (
                <Error
                  icon={LoadingMessagesErrorIcon}
                  message="No messages were found for this chat."
                  subheading={TypeANewMessage}
                ></Error>
              )}
              {chatState.status === 'no chat id' && (
                <Error message="No chat id has been provided." subheading={RequireValidChatId}></Error>
              )}
              {chatState.status === 'error' && (
                <Error message="We're sorry—we've run into an issue.." subheading={OpenTeamsLinkError}></Error>
              )}
              <div className={styles.chatInput}>
                <SendBox disabled={disabled} onSendMessage={chatState.onSendMessage} strings={{ placeholderText }} />
              </div>
            </>
          )}
        </div>
      </FluentProvider>
    </FluentThemeProvider>
  );
};
