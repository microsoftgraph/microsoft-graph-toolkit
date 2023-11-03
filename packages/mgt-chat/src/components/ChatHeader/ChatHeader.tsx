import { Divider, makeStyles, mergeClasses } from '@fluentui/react-components';
import React from 'react';
import { GraphChatClient } from '../../statefulClient/StatefulGraphChatClient';
import ChatTitle from './ChatTitle';
import { ManageChatMembers } from '../ManageChatMembers/ManageChatMembers';
import { useCommonHeaderStyles } from './useCommonHeaderStyles';
import { EllipsisMenu } from './EllipsisMenu';

const useHeaderStyles = makeStyles({
  chatHeader: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: 0,
    zIndex: 2
  },
  secondRow: {
    display: 'flex',
    flexDirection: 'row',
    justifyContent: 'flex-start',
    alignItems: 'center',
    height: '30px'
  },
  noGrow: {
    flexGrow: 0
  }
});
interface ChatHeaderProps {
  chatState: GraphChatClient;
}
const ChatHeader = ({ chatState }: ChatHeaderProps) => {
  const styles = useHeaderStyles();
  const commonStyles = useCommonHeaderStyles();
  const chatWebUrl = chatState?.chat?.webUrl ?? 'https://teams.microsoft.com/v2';

  return (
    <div className={styles.chatHeader}>
      <ChatTitle chat={chatState.chat} currentUserId={chatState.userId} onRenameChat={chatState.onRenameChat} />
      <Divider appearance="subtle" />
      <div className={mergeClasses(styles.secondRow, commonStyles.row)}>
        {chatState.participants?.length > 0 && chatState.chat?.chatType === 'group' && (
          <>
            <ManageChatMembers
              members={chatState.participants}
              removeChatMember={chatState.onRemoveChatMember}
              currentUserId={chatState.userId}
              addChatMembers={chatState.onAddChatMembers}
            />
            <Divider vertical appearance="subtle" inset className={styles.noGrow} />
          </>
        )}
        <EllipsisMenu chatWebUrl={chatWebUrl} />
      </div>
      <Divider appearance="subtle" />
    </div>
  );
};

export { ChatHeader };
