import React, { useEffect, useState, useCallback } from 'react';
import { ChatListItem, IChatListItemInteractionProps } from '../ChatListItem/ChatListItem';
import { MgtTemplateProps, ProviderState, Providers } from '@microsoft/mgt-react';
import { makeStyles, Link, FluentProvider, shorthands, webLightTheme } from '@fluentui/react-components';
import { FluentThemeProvider } from '@azure/communication-react';
import { FluentTheme } from '@fluentui/react';
import { Chat as GraphChat } from '@microsoft/microsoft-graph-types';
import { ChatListEvent, StatefulGraphChatListClient } from '../../statefulClient/StatefulGraphChatListClient';
import { useGraphChatListClient } from '../../statefulClient/useGraphChatListClient';
import { ChatListHeader } from '../ChatListHeader/ChatListHeader';
import { IChatListMenuItemsProps } from '../ChatListHeader/EllipsisMenu';
import { ChatListButtonItem } from '../ChatListHeader/ChatListButtonItem';
import { ChatThreadCollection, loadChatThreads, loadChatThreadsByPage } from '../../statefulClient/graph.chat';
import { error } from '@microsoft/mgt-element';

const useStyles = makeStyles({
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
  loadMore: {
    textDecorationLine: 'none',
    fontSize: '1.2em',
    fontWeight: 'bold',
    '&:hover': {
      textDecorationLine: 'none' // This removes the underline when hovering
    }
  }
});

// this is a stub to move the logic here that should end up here.
export const ChatList = (
  props: MgtTemplateProps &
    IChatListItemInteractionProps &
    IChatListMenuItemsProps & {
      buttonItems?: ChatListButtonItem[];
      chatThreadsPerPage: number;
    }
) => {
  const styles = useStyles();
  const chatClient: StatefulGraphChatListClient = useGraphChatListClient();
  const [chatState, setChatState] = useState(chatClient.getState());
  const [chatThreads, setChatThreads] = useState<GraphChat[]>([]);
  const [nextLink, setNextLink] = useState<string>('');

  const loadMore = () => {
    if (nextLink && nextLink !== '') {
      const provider = Providers.globalProvider;
      if (provider && provider.state === ProviderState.SignedIn) {
        let filter = nextLink.split('?')[1];
        setNextLink(''); // reset

        const graph = provider.graph.forComponent('ChatList');
        loadChatThreadsByPage(graph, filter).then(
          chats => handleChatThreads(chats),
          e => error(e)
        );
      }
    }
  };

  const handleChatThreads = (chatThreadCollection: ChatThreadCollection) => {
    let nextLinkUrl = chatThreadCollection['@odata.nextLink'];
    if (nextLinkUrl && nextLinkUrl !== '') {
      setNextLink(nextLinkUrl);
    }

    console.log('chatThreadCollection:' + chatThreadCollection.value);

    let uniqeChatThreads = chatThreadCollection.value.filter(c => chatThreads.findIndex(t => t.id === c.id) === -1);
    setChatThreads(chatThreads.concat(uniqeChatThreads));
  };

  const loadDataOnlyOnce = useCallback(() => {
    const provider = Providers.globalProvider;
    if (provider && provider.state === ProviderState.SignedIn) {
      const graph = provider.graph.forComponent('ChatList');
      loadChatThreads(graph, props.chatThreadsPerPage).then(
        chats => handleChatThreads(chats),
        e => error(e)
      );
    }
  }, [chatThreads]);

  useEffect(loadDataOnlyOnce, []);

  const onChatListEvent = (state: ChatListEvent) => {
    // TODO: implementation will happen later, right now, we just need to make sure messages are coming thru in console logs.

    console.log(state.type);
    console.log(state.message);
  };

  useEffect(() => {
    chatClient.onChatListEvent(onChatListEvent);
    chatClient.onStateChange(setChatState);
    return () => {
      chatClient.offChatListEvent(onChatListEvent);
      chatClient.offStateChange(setChatState);
    };
  }, [chatClient]);

  const chatListButtonItems = props.buttonItems === undefined ? [] : props.buttonItems;
  const chatListMenuItems = props.menuItems === undefined ? [] : props.menuItems;
  chatListMenuItems.unshift({
    displayText: 'Mark all as read',
    onClick: () => {}
  });

  return (
    // This is a temporary approach to render the chatlist items. This should be replaced.
    <FluentThemeProvider fluentTheme={FluentTheme}>
      <FluentProvider theme={webLightTheme}>
        <div>
          <div className={styles.headerContainer}>
            <ChatListHeader buttonItems={chatListButtonItems} menuItems={chatListMenuItems} />
          </div>
          {chatThreads.map(c => (
            <ChatListItem key={c.id} chat={c} myId={chatState.userId} onSelected={props.onSelected} />
          ))}
          {nextLink !== '' && (
            <div className={styles.linkContainer}>
              <Link onClick={loadMore} href="#" className={styles.loadMore}>
                load more
              </Link>
            </div>
          )}
        </div>
      </FluentProvider>
    </FluentThemeProvider>
  );
};

export default ChatList;
