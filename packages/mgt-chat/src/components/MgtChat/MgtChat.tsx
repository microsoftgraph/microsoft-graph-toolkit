import React, { useEffect, useState } from 'react';
import {
  ErrorBar,
  FluentThemeProvider,
  MessageThread,
  SendBox,
  DEFAULT_COMPONENT_ICONS
} from '@azure/communication-react';
import { registerIcons } from '@fluentui/react';
import { Person, PersonCardInteraction, Spinner } from '@microsoft/mgt-react';
import { useGraphChatClient } from '../../statefulClient/useGraphChatClient';
import ChatHeader from '../ChatHeader/ChatHeader';

import * as styles from './chat.module.scss';

registerIcons({ icons: DEFAULT_COMPONENT_ICONS });
interface IMgtChatProps {
  chatId: string;
}

export const MgtChat = ({ chatId }: IMgtChatProps) => {
  const chatClient = useGraphChatClient(chatId);
  const [chatState, setChatState] = useState(chatClient.getState());
  useEffect(() => {
    chatClient.onStateChange(setChatState);
    return () => {
      chatClient.offStateChange(setChatState);
    };
  }, [setChatState]);
  return (
    <FluentThemeProvider>
      <div className={styles.chat}>
        {chatState.userId && chatState.messages.length > 0 ? (
          <>
            <ChatHeader />
            <div className={styles.chat__messages}>
              <MessageThread
                userId={chatState.userId}
                messages={chatState.messages}
                showMessageDate={true}
                disableEditing={chatState.disableEditing}
                // TODO: Establish how to stop loading more messages and have it load more as part of an infinite scroll
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
            <div className={styles.chat__input}>
              <SendBox autoFocus="sendBoxTextField" onSendMessage={chatState.onSendMessage} />
            </div>
            <ErrorBar activeErrorMessages={chatState.activeErrorMessages} />
          </>
        ) : (
          <>
            {chatState.status}
            <Spinner />
          </>
        )}
      </div>
    </FluentThemeProvider>
  );
};
