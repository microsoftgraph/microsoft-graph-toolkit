import {
  ErrorBar,
  FluentThemeProvider,
  MessageProps,
  MessageRenderer,
  MessageThread,
  SendBox
} from '@azure/communication-react';
<<<<<<< HEAD
import { Person, PersonCardInteraction, Spinner } from '@microsoft/mgt-react';
import { FluentTheme, MessageBarType } from '@fluentui/react';
=======
import { FluentTheme } from '@fluentui/react';
>>>>>>> 0397e88d (Update the Chat to use variables from the state object)
import { FluentProvider, makeStyles, shorthands, teamsLightTheme } from '@fluentui/react-components';
import { Person, PersonCardInteraction, Spinner } from '@microsoft/mgt-react';
import React, { useEffect, useState } from 'react';
import { renderToString } from 'react-dom/server';
import { useGraphChatClient } from '../../statefulClient/useGraphChatClient';
import { isChatMessage, isGraphChatMessage } from '../../utils/types';
import ChatHeader from '../ChatHeader/ChatHeader';
<<<<<<< HEAD
import ChatMessageBar from '../ChatMessageBar/ChatMessageBar';
import { registerAppIcons } from '../styles/registerIcons';
=======
>>>>>>> 0397e88d (Update the Chat to use variables from the state object)
import { ManageChatMembers } from '../ManageChatMembers/ManageChatMembers';
import { StatefulGraphChatClient } from 'src/statefulClient/StatefulGraphChatClient';
import UnsupportedContent from '../UnsupportedContent/UnsupportedContent';
import { registerAppIcons } from '../styles/registerIcons';

registerAppIcons();

interface IMgtChatProps {
  chatId: string;
}

const useStyles = makeStyles({
  chat: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    ...shorthands.overflow('auto')
  },
  chatMessages: {
    height: 'auto',
    ...shorthands.overflow('auto'),
    '& img': {
      maxWidth: '100%',
      height: 'auto'
    }
  },
  chatInput: {
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
 
  const onRenderMessage = (messageProps: MessageProps, defaultOnRender?: MessageRenderer) => {
    const updatedProps = Object.assign({}, { ...messageProps });
    const message = updatedProps?.message;
    if (isGraphChatMessage(message) && message?.hasUnsupportedContent) {
      const unsupportedContentComponent = <UnsupportedContent targetUrl={message.rawChatUrl} />;
      if (isChatMessage(message)) {
        // TODO: assigning this string to content fails because props are
        // TODO: readonly. Re-introduce produce?
        message.content = renderToString(unsupportedContentComponent);
      }
    }

    return defaultOnRender ? defaultOnRender(updatedProps) : <></>;
  };

  return (
    <FluentThemeProvider fluentTheme={FluentTheme}>
      <FluentProvider theme={teamsLightTheme} className={styles.fullHeight}>
        <div className={styles.chat}>
          {chatState.userId && chatId && chatState.messages.length > 0 ? (
            <>
              <ChatHeader
                chat={chatState.chat}
                currentUserId={chatState.userId}
                onRenameChat={chatState.onRenameChat}
              />
              {chatState.participants?.length > 0 && chatState.chat?.chatType === 'group' && (
                <ManageChatMembers
                  members={chatState.participants}
                  removeChatMember={chatState.onRemoveChatMember}
                  currentUserId={chatState.userId}
                  addChatMembers={chatState.onAddChatMembers}
                />
              )}
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
                      <Person userId={userId} avatarSize="small" personCardInteraction={PersonCardInteraction.click} />
                    );
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
