import { ErrorBar, FluentThemeProvider, MessageThread, SendBox, MessageThreadStyles } from '@azure/communication-react';
import { FluentTheme, MessageBarType } from '@fluentui/react';
import { FluentProvider, makeStyles, shorthands, webLightTheme } from '@fluentui/react-components';
import { Person, PersonCardInteraction, Spinner } from '@microsoft/mgt-react';
import React, { useEffect, useState } from 'react';
import { StatefulGraphChatClient } from '../../statefulClient/StatefulGraphChatClient';
import { useGraphChatClient } from '../../statefulClient/useGraphChatClient';
import { onRenderMessage } from '../../utils/chat';
import ChatMessageBar from '../ChatMessageBar/ChatMessageBar';
import { renderMGTMention } from '../../utils/mentions';
import { registerAppIcons } from '../styles/registerIcons';
import { ChatHeader } from '../ChatHeader/ChatHeader';

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
    paddingBlockEnd: '12px'
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
      zIndex: 'unset'
    },
    '& p': {
      display: 'inline-flex',
      justifyContent: 'center'
    }
  },
  chatMessageContainer: {
    '& p>mgt-person,msft-mention': {
      display: 'inline-block',
      ...shorthands.marginInline('0px', '2px')
    },
    '& .otherMention': {
      color: 'var(--accent-base-color)'
    }
  },
  myChatMessageContainer: {
    '& p>mgt-person,msft-mention': {
      display: 'inline-block',
      ...shorthands.marginInline('0px', '2px')
    },
    '& .otherMention': {
      color: 'var(--accent-base-color)'
    }
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

  return (
    <FluentThemeProvider fluentTheme={FluentTheme}>
      <FluentProvider theme={webLightTheme} className={styles.fullHeight}>
        <div className={styles.chat}>
          {chatState.userId && chatId && chatState.messages.length > 0 ? (
            <>
              <ChatHeader chatState={chatState} />
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
                      <Person
                        userId={userId}
                        avatarSize="small"
                        personCardInteraction={PersonCardInteraction.hover}
                        showPresence={true}
                      />
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
                <SendBox onSendMessage={chatState.onSendMessage} />
              </div>
              <ErrorBar activeErrorMessages={chatState.activeErrorMessages} />
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
                <ChatMessageBar
                  messageBarType={MessageBarType.error}
                  message={`No messages were found for the id ${chatId}.`}
                />
              )}
              {chatState.status === 'no chat id' && (
                <ChatMessageBar messageBarType={MessageBarType.error} message={'A valid chat id is required.'} />
              )}
            </>
          )}
        </div>
      </FluentProvider>
    </FluentThemeProvider>
  );
};
