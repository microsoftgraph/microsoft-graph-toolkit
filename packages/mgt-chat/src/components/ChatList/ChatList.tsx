import React, { useEffect, useState, useRef } from 'react';
import { ChatListItem } from '../ChatListItem/ChatListItem';
import { MgtTemplateProps, ProviderState, Providers, Spinner, log } from '@microsoft/mgt-react';
import { makeStyles, Button, FluentProvider, shorthands, webLightTheme } from '@fluentui/react-components';
import { FluentThemeProvider } from '@azure/communication-react';
import { FluentTheme } from '@fluentui/react';
import { ChatMessage } from '@microsoft/microsoft-graph-types';
import {
  StatefulGraphChatListClient,
  GraphChatListClient,
  GraphChatThread
} from '../../statefulClient/StatefulGraphChatListClient';
import { ChatListHeader } from '../ChatListHeader/ChatListHeader';
import { IChatListMenuItemsProps } from '../ChatListHeader/EllipsisMenu';
import { ChatListButtonItem } from '../ChatListHeader/ChatListButtonItem';
import { Error } from '../Error/Error';
import { LoadingMessagesErrorIcon } from '../Error/LoadingMessageErrorIcon';
import { CreateANewChat } from '../Error/CreateANewChat';
import { PleaseSignIn } from '../Error/PleaseSignIn';
import { OpenTeamsLinkError } from '../Error/OpenTeams';
import IChatListActions from '../ChatListHeader/IChatListActions';

export interface IChatListProps {
  onSelected: (e: GraphChatThread) => void;
  onUnselected?: (e: GraphChatThread) => void;
  onLoaded?: (e: GraphChatThread[]) => void;
  onAllMessagesRead: (e: string[]) => void;
  buttonItems?: ChatListButtonItem[];
  chatThreadsPerPage: number;
  lastReadTimeInterval?: number;
  selectedChatId?: string;
  onMessageReceived?: (msg: ChatMessage) => void;
  onConnectionChanged?: (connected: boolean) => void;
}

const useStyles = makeStyles({
  chatList: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    ...shorthands.overflow('hidden'),
    paddingBlockEnd: '12px'
  },
  chatListItems: {
    height: '100%',
    ...shorthands.paddingInline('20px'),
    '&': {
      paddingRight: '8px' // reserved some space for the scrollbar
    },
    '&:hover': {
      paddingRight: '0', // we got the scrollbar, no need to reserve space
      ...shorthands.overflow('auto'),
      scrollbarWidth: 'auto'
    },
    '&:hover::-webkit-scrollbar': {
      width: '8px'
    },
    '&:hover::-webkit-scrollbar-thumb': {
      backgroundColor: 'darkgrey',
      'border-radius': '5px'
    }
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
  button: {
    flexDirection: 'row',
    alignItems: 'center',
    width: '100%',
    ...shorthands.padding('0px'),
    ...shorthands.border('none')
  },
  error: {
    display: 'flex',
    justifyContent: 'center',
    height: '100%'
  },
  bottomWhitespace: {
    height: '80%',
    width: '100%'
  }
});

// this is a stub to move the logic here that should end up here.
export const ChatList = ({
  lastReadTimeInterval = 30000, // default to 30 seconds
  selectedChatId,
  onMessageReceived,
  onAllMessagesRead,
  onLoaded,
  onConnectionChanged,
  onSelected,
  onUnselected,
  chatThreadsPerPage,
  ...props
}: MgtTemplateProps & IChatListProps & IChatListMenuItemsProps) => {
  const styles = useStyles();

  const [chatListClient, setChatListClient] = useState<StatefulGraphChatListClient | undefined>();
  const [chatListState, setChatListState] = useState<GraphChatListClient | undefined>();
  const [chatListActions, setChatListActions] = useState<IChatListActions | undefined>();
  const loadingRef = useRef(false);
  // wait for provider to be ready before setting client and state
  useEffect(() => {
    const provider = Providers.globalProvider;
    const conditionalLoad = (state: ProviderState | undefined) => {
      if (state === ProviderState.SignedIn && !chatListClient) {
        const client = new StatefulGraphChatListClient(chatThreadsPerPage, selectedChatId);
        setChatListClient(client);
        setChatListState(client.getState());
      }
    };
    provider?.onStateChanged(evt => {
      conditionalLoad(evt.detail);
    });
    conditionalLoad(provider?.state);
  }, [chatListClient, chatThreadsPerPage, selectedChatId]);

  useEffect(() => {
    if (chatListClient) {
      setChatListActions({
        markAllChatThreadsAsRead: () => chatListClient.markAllChatThreadsAsRead()
      });
    }
  }, [chatListClient]);

  // Store last read time in cache so that when the user comes back to the chat list,
  // we know what messages they are likely to have not read. This is not perfect because
  // the user could have read messages in another client (for instance, the Teams client).
  useEffect(() => {
    const timer = setInterval(() => {
      chatListClient?.cacheLastReadTime('selected');
    }, lastReadTimeInterval);

    return () => {
      clearInterval(timer);
    };
  }, [chatListClient, lastReadTimeInterval]);

  useEffect(() => {
    // shortcut if we don't have a chat list client
    if (!chatListClient) {
      return;
    }

    // handle state changes
    chatListClient.onStateChange(setChatListState);
    chatListClient.onStateChange(state => {
      if (state.status === 'chat message received' && onMessageReceived && state.chatMessage) {
        onMessageReceived(state.chatMessage);
      }

      if (state.status === 'chat selected' && onSelected && state.internalSelectedChat) {
        onSelected(state.internalSelectedChat);
      }

      if (state.status === 'chat unselected' && onUnselected && state.internalPrevSelectedChat) {
        onUnselected(state.internalPrevSelectedChat);
      }

      if (state.status === 'chats read' && onAllMessagesRead && state.chatThreads) {
        onAllMessagesRead(state.chatThreads.map(c => c.id!));
      }

      if (state.status === 'chats loaded' && onLoaded) {
        onLoaded(state?.chatThreads ?? []);
        // the loadingRef is used to prevent multiple calls to loadMoreChatThreads
        loadingRef.current = false;
      }

      if (state.status === 'chats loaded' && state.initialSelectedChatSet && onSelected && state.internalSelectedChat) {
        onSelected(state.internalSelectedChat);
      }

      if (state.status === 'no chats' && onLoaded) {
        onLoaded([]);
      }

      if (state.status === 'server connection established' && onConnectionChanged) {
        onConnectionChanged(true);
      }

      if (state.status === 'server connection lost' && onConnectionChanged) {
        onConnectionChanged(false);
      }
    });
  }, [chatListClient, onLoaded, onMessageReceived, onSelected, onUnselected, onAllMessagesRead, onConnectionChanged]);

  // this only runs once when the component is unmounted
  useEffect(() => {
    if (chatListClient) {
      // tear down
      return () => {
        // log state of chatlistclient for debugging purposes
        log('ChatList unmounted.', chatListClient.getState());
        chatListClient.offStateChange(setChatListState);
        chatListClient.tearDown();
      };
    }
  }, []);

  const onClickChatListItem = (chat: GraphChatThread) => {
    chatListClient?.setInternalSelectedChat(chat);
  };

  const chatListButtonItems = props.buttonItems === undefined ? [] : props.buttonItems;
  const chatListMenuItems = props.menuItems === undefined ? [] : props.menuItems;

  const isLoading = ['creating server connections', 'subscribing to notifications', 'loading messages'].includes(
    chatListState?.status ?? ''
  );

  const targetElementRef = useRef(null);

  useEffect(() => {
    const handleIntersection = (entries: IntersectionObserverEntry[]) => {
      for (const entry of entries) {
        if (entry.isIntersecting) {
          // The element has come into view, you can perform your actions here
          if (chatListClient && !loadingRef.current) {
            // Prevent the function from being called multiple times
            if (chatListState?.moreChatThreadsToLoad) {
              void chatListClient.loadMoreChatThreads().then(() => (loadingRef.current = true));
            }
          }
        }
      }
    };
    // Create a new Intersection Observer instance
    const observer = new IntersectionObserver(handleIntersection, {
      root: null, // observing intersections with the viewport
      rootMargin: '0px',
      threshold: 0.1 // Callback is invoked when 10% of the target is visible
    });

    if (targetElementRef.current) {
      observer.observe(targetElementRef.current);
    }

    return () => observer.disconnect();
  }, [chatListState]);

  return (
    <FluentThemeProvider fluentTheme={FluentTheme}>
      <FluentProvider theme={webLightTheme} className={styles.fullHeight}>
        <div className={styles.chatList}>
          {Providers.globalProvider?.state === ProviderState.SignedIn &&
            chatListState?.status !== 'server connection lost' &&
            chatListActions && (
              <ChatListHeader
                actions={chatListActions}
                buttonItems={chatListButtonItems}
                menuItems={chatListMenuItems}
              />
            )}
          {chatListState && chatListState.chatThreads.length > 0 ? (
            <>
              <div className={styles.chatListItems}>
                {chatListState?.chatThreads.map(c => (
                  <Button className={styles.button} key={c.id} onClick={() => onClickChatListItem(c)}>
                    <ChatListItem
                      key={c.id}
                      chat={c}
                      myId={chatListState.userId}
                      isSelected={c.id === chatListState?.internalSelectedChat?.id}
                      isRead={c.isRead}
                    />
                  </Button>
                ))}
                {chatListState?.moreChatThreadsToLoad && (
                  <div ref={targetElementRef} className={styles.bottomWhitespace}>
                    &nbsp;
                  </div>
                )}
              </div>
            </>
          ) : (
            <>
              <div className={styles.error}>
                {!chatListState?.userId && <Error message="User not signed-in." subheading={PleaseSignIn}></Error>}
                {isLoading && (
                  <div className={styles.spinner}>
                    <Spinner /> <br />
                    {chatListState?.status}
                  </div>
                )}
                {chatListState?.status === 'no chats' && (
                  <Error
                    icon={LoadingMessagesErrorIcon}
                    message="No threads were found for this user."
                    subheading={CreateANewChat}
                  ></Error>
                )}
                {chatListState?.status === 'server connection lost' && (
                  <Error message="We ran into a problem. Reconnecting..." subheading={OpenTeamsLinkError}></Error>
                )}
              </div>
            </>
          )}
        </div>
      </FluentProvider>
    </FluentThemeProvider>
  );
};
