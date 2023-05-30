import React, { useEffect, useState } from 'react';
import { CallingTheme, ErrorBar, FluentThemeProvider, MessageThread, SendBox } from '@azure/communication-react';
import { Person, PersonCardInteraction, Spinner } from '@microsoft/mgt-react';
import { FluentTheme, PartialTheme } from '@fluentui/react';
import { Divider, FluentProvider, Theme, makeStyles, shorthands, teamsLightTheme } from '@fluentui/react-components';
import { useGraphChatClient } from '../../statefulClient/useGraphChatClient';
import ChatHeader from '../ChatHeader/ChatHeader';
import { registerAppIcons } from '../styles/registerIcons';
import { ManageChatMembers } from '../ManageChatMembers/ManageChatMembers';
import { chatStyles } from './chat.styles';

registerAppIcons();

interface IMgtChatProps {
  chatId: string;
  chatTheme?: PartialTheme & CallingTheme;
  fluentTheme?: Theme;
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
  }
});

export const Chat = ({ chatId, chatTheme, fluentTheme }: IMgtChatProps) => {
  const styles = useStyles();
  const chatClient = useGraphChatClient(chatId);
  const [chatState, setChatState] = useState(chatClient.getState());
  useEffect(() => {
    chatClient.onStateChange(setChatState);
    return () => {
      chatClient.offStateChange(setChatState);
    };
  }, [chatClient]);
  return (
    <FluentThemeProvider fluentTheme={chatTheme ? chatTheme : FluentTheme}>
      <FluentProvider theme={fluentTheme ? fluentTheme : teamsLightTheme} className={chatStyles.fullHeight}>
        <div className={chatStyles.chat}>
          {chatState.userId && chatState.messages.length > 0 ? (
            <>
              <ChatHeader
                chat={chatState.chat}
                currentUserId={chatState.userId}
                onRenameChat={chatState.onRenameChat}
              />
              {chatState.participants?.length > 0 && chatState.chat?.chatType === 'group' && (
                <>
                  <Divider></Divider>
                  <ManageChatMembers
                    members={chatState.participants}
                    removeChatMember={chatState.onRemoveChatMember}
                    currentUserId={chatState.userId}
                    addChatMembers={chatState.onAddChatMembers}
                  />
                  <Divider></Divider>
                </>
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
                />
              </div>
              <div className={styles.chatInput}>
                <SendBox onSendMessage={chatState.onSendMessage} />
              </div>
              <ErrorBar activeErrorMessages={chatState.activeErrorMessages} />
            </>
          ) : (
            <div className={chatStyles.spinnerContainer}>
              {/* chatState.status */}
              <Spinner />
            </div>
          )}
        </div>
      </FluentProvider>
    </FluentThemeProvider>
  );
};
