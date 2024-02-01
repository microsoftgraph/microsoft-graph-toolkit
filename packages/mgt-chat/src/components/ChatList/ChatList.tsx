import React, { useEffect, useState } from 'react';
import { ChatListItem } from '../ChatListItem/ChatListItem';
import { MgtTemplateProps, ProviderState, Providers, Spinner, log } from '@microsoft/mgt-react';
import { makeStyles, Button, Link, FluentProvider, shorthands, webLightTheme } from '@fluentui/react-components';
import { FluentThemeProvider } from '@azure/communication-react';
import { FluentTheme } from '@fluentui/react';
import { Chat as GraphChat, ChatMessage } from '@microsoft/microsoft-graph-types';
import {
  StatefulGraphChatListClient,
  GraphChatListClient,
  ChatListEvent
} from '../../statefulClient/StatefulGraphChatListClient';
import { ChatListHeader } from '../ChatListHeader/ChatListHeader';
import { IChatListMenuItemsProps } from '../ChatListHeader/EllipsisMenu';
import { ChatListButtonItem } from '../ChatListHeader/ChatListButtonItem';
import { Error } from '../Error/Error';
import { LoadingMessagesErrorIcon } from '../Error/LoadingMessageErrorIcon';
import { CreateANewChat } from '../Error/CreateANewChat';
import { PleaseSignIn } from '../Error/PleaseSignIn';
import { OpenTeamsLinkError } from '../Error/OpenTeams';

export interface IChatListItemProps {
  onSelected: (e: GraphChat) => void;
  onUnselected?: (e: GraphChat) => void;
  onLoaded?: () => void;
  onAllMessagesRead: (e: string[]) => void;
  buttonItems?: ChatListButtonItem[];
  chatThreadsPerPage: number;
  lastReadTimeInterval?: number;
  selectedChatId?: string;
  onMessageReceived?: (msg: ChatMessage) => void;
}

const useStyles = makeStyles({
  fullHeight: {
    height: '100%'
  },
  headerContainer: {
    display: 'flex',
    justifyContent: 'center',
    width: '100%',
    ...shorthands.padding('10px')
  },
  linkContainer: {
    display: 'flex',
    justifyContent: 'center',
    width: '100%',
    ...shorthands.padding('10px')
  },
  spinner: {
    justifyContent: 'center',
    display: 'flex',
    alignItems: 'center',
    height: '100%'
  },
  loadMore: {
    textDecorationLine: 'none',
    fontSize: '1.2em',
    fontWeight: 'bold',
    '&:hover': {
      textDecorationLine: 'none' // This removes the underline when hovering
    }
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
  }
});

// this is a stub to move the logic here that should end up here.
export const ChatList = ({
  lastReadTimeInterval = 30000, // default to 30 seconds
  selectedChatId,
  onMessageReceived,
  onAllMessagesRead,
  onLoaded,
  chatThreadsPerPage,
  ...props
}: MgtTemplateProps & IChatListItemProps & IChatListMenuItemsProps) => {
  const styles = useStyles();
  const [headerBannerMessage, setHeaderBannerMessage] = useState<string>('');
  const [chatListClient, setChatListClient] = useState<StatefulGraphChatListClient | undefined>();
  const [chatListState, setChatListState] = useState<GraphChatListClient | undefined>();
  const [internalSelectedChatId, setInternalSelectedChatId] = useState<string | undefined>();

  // wait for provider to be ready before setting client and state
  useEffect(() => {
    const provider = Providers.globalProvider;
    const conditionalLoad = (state: ProviderState | undefined) => {
      if (state === ProviderState.SignedIn && !chatListClient) {
        const client = new StatefulGraphChatListClient(chatThreadsPerPage);
        setChatListClient(client);
        setChatListState(client.getState());
      }
    };
    provider.onStateChanged(evt => {
      conditionalLoad(evt.detail);
    });
    conditionalLoad(provider?.state);
  }, [chatListClient, chatThreadsPerPage]);

  // if selected chat id is changed, update the internal state
  // NOTE: Decoupling this ensures that the app can change the selection but the chat
  // list can also change the selection without a full re - render.
  useEffect(() => {
    setInternalSelectedChatId(selectedChatId);
  }, [selectedChatId]);

  // Store last read time in cache so that when the user comes back to the chat list,
  // we know what messages they are likely to have not read. This is not perfect because
  // the user could have read messages in another client (for instance, the Teams client).
  useEffect(() => {
    const timer = setInterval(() => {
      if (internalSelectedChatId) {
        log(`caching the last-read timestamp of now to chat ID '${internalSelectedChatId}'...`);
        chatListClient?.cacheLastReadTime([internalSelectedChatId]);
      }
    }, lastReadTimeInterval);

    return () => {
      clearInterval(timer);
    };
  }, [chatListClient, internalSelectedChatId, lastReadTimeInterval]);

  useEffect(() => {
    // shortcut if we don't have a chat list client
    if (!chatListClient) {
      return;
    }

    // handle events emitted from the chat list client
    const handleChatListEvent = (event: ChatListEvent) => {
      if (event.type === 'chatMessageReceived') {
        if (onMessageReceived) {
          onMessageReceived(event.message);
        }
      }
    };

    // handle state changes
    chatListClient.onStateChange(setChatListState);
    chatListClient.onStateChange(state => {
      if (state.status === 'chats loaded' && onLoaded) {
        onLoaded();
      }

      if (state.status === 'server connection established') {
        setHeaderBannerMessage(''); // reset
      }

      if (state.status === 'creating server connections') {
        setHeaderBannerMessage('Connecting...');
      }

      if (state.status === 'server connection lost') {
        // this happens when we lost connection to the server and we will try to reconnect
        setHeaderBannerMessage('We ran into a problem. Reconnecting...');
      }
    });

    // handle chat list events
    chatListClient.onChatListEvent(handleChatListEvent);

    // tear down
    return () => {
      // log state of chatlistclient for debugging purposes
      log(chatListClient.getState());
      chatListClient.offStateChange(setChatListState);
      chatListClient.offChatListEvent(handleChatListEvent);
      chatListClient.tearDown();
      setHeaderBannerMessage('We ran into a problem. Please close or refresh.');
    };
  }, [chatListClient, onMessageReceived, onLoaded]);

  const markThreadAsRead = (chatThread: string) => {
    const markedChatThreads = chatListClient?.markChatThreadsAsRead([chatThread]);
    if (markedChatThreads) {
      chatListClient?.cacheLastReadTime(markedChatThreads);
    }
  };

  const onClickChatListItem = (chatListItem: GraphChat) => {
    // set selected state only once per click event
    if (chatListItem.id !== internalSelectedChatId) {
      // trigger an unselect event for the previously selected item
      if (internalSelectedChatId && props.onUnselected) {
        const previouslySelectedChatListItem = chatListState?.chatThreads.filter(c => c.id === internalSelectedChatId);
        if (previouslySelectedChatListItem?.length === 1) {
          props.onUnselected(previouslySelectedChatListItem[0]);
        }
      }

      // select a new item
      markThreadAsRead(chatListItem.id!);
      setInternalSelectedChatId(chatListItem.id);
      props.onSelected(chatListItem);
    }
  };

  const chatListButtonItems = props.buttonItems === undefined ? [] : props.buttonItems;

  // We need to have a function for "this" to work within the loadMoreChatThreads function, otherwise we get a undefined error.
  const loadMore = () => {
    chatListClient?.loadMoreChatThreads();
  };

  const markAllThreadsAsRead = (chatThreads: GraphChat[] | undefined) => {
    if (!chatThreads) {
      return;
    }
    const readChatThreads = chatThreads.map(c => c.id!);
    const markedChatThreads = chatListClient?.markChatThreadsAsRead(readChatThreads);
    if (markedChatThreads) {
      chatListClient?.cacheLastReadTime(markedChatThreads);
      onAllMessagesRead(markedChatThreads);
    }
  };
  const markAllAsRead = {
    displayText: 'Mark all as read',
    onClick: () => markAllThreadsAsRead(chatListState?.chatThreads)
  };

  const isLoading = [
    'creating server connections',
    'server connection established',
    'subscribing to notifications',
    'loading messages'
  ].includes(chatListState?.status ?? '');

  const isError = ['server connection lost', 'error'].includes(chatListState?.status ?? '');

  return (
    <FluentThemeProvider fluentTheme={FluentTheme}>
      <FluentProvider theme={webLightTheme} className={styles.fullHeight}>
        <div className={styles.fullHeight}>
          {Providers.globalProvider?.state === ProviderState.SignedIn && (
            <div className={styles.headerContainer}>
              <ChatListHeader
                bannerMessage={headerBannerMessage}
                buttonItems={chatListButtonItems}
                menuItems={[markAllAsRead, ...(props.menuItems ?? [])]}
              />
            </div>
          )}
          {chatListState && chatListState.chatThreads.length > 0 ? (
            <>
              <div>
                {chatListState?.chatThreads.map(c => (
                  <Button className={styles.button} key={c.id} onClick={() => onClickChatListItem(c)}>
                    <ChatListItem
                      key={c.id}
                      chat={c}
                      myId={chatListState.userId}
                      isSelected={c.id === internalSelectedChatId}
                      isRead={c.id === internalSelectedChatId || c.isRead}
                    />
                  </Button>
                ))}
                {chatListState?.moreChatThreadsToLoad === true && (
                  <div className={styles.linkContainer}>
                    <Link onClick={loadMore} href="#" className={styles.loadMore}>
                      load more
                    </Link>
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
                {isError && (
                  <Error message="We're sorry—we've run into an issue." subheading={OpenTeamsLinkError}></Error>
                )}
              </div>
            </>
          )}
        </div>
      </FluentProvider>
    </FluentThemeProvider>
  );
};
