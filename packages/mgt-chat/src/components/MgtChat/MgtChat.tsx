import React, { useEffect, useState } from 'react';
import { ErrorBar, FluentThemeProvider, MessageThread, SendBox } from '@azure/communication-react';
import { useGraphChatClient } from '../../statefulClient/useGraphChatClient';
import { registerIcons } from '@fluentui/react';
import { DEFAULT_COMPONENT_ICONS } from '@azure/communication-react';
registerIcons({ icons: DEFAULT_COMPONENT_ICONS });
interface IMgtChatProps {
  chatId: string;
}

const MgtChat = ({ chatId }: IMgtChatProps) => {
  const chatClient = useGraphChatClient(chatId);
  const [chatState, setChatState] = useState(chatClient.getState());
  useEffect(() => {
    chatClient.onStateChange(setChatState);
    return () => {
      chatClient.offStateChange(setChatState);
    };
  }, [setChatState]);
  return (
    <>
      Rendering chat: {chatId}
      {chatState.userId && chatState.messages && (
        <FluentThemeProvider>
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
            onSendMessage={chatState.onResendMessage}
          />
          <SendBox autoFocus="sendBoxTextField" onSendMessage={chatState.onSendMessage} />
          <ErrorBar activeErrorMessages={chatState.activeErrorMessages} />
        </FluentThemeProvider>
      )}
    </>
  );
};

export default MgtChat;
