import React, { memo, useCallback, useState } from 'react';
import { AadUserConversationMember, Chat } from '@microsoft/microsoft-graph-types';
import { Edit20Filled, Edit20Regular, bundleIcon } from '@fluentui/react-icons';
import { Person, PersonCardInteraction, ViewType } from '@microsoft/mgt-react';
import {
  Button,
  Field,
  Input,
  InputOnChangeData,
  OnOpenChangeData,
  Popover,
  PopoverSurface,
  PopoverTrigger,
  makeStyles,
  tokens
} from '@fluentui/react-components';

interface ChatHeaderProps {
  chat?: Chat;
  currentUserId?: string;
  onRenameChat: (newName: string | null) => Promise<void>;
}

const EditIcon = bundleIcon(Edit20Filled, Edit20Regular);

const useStyles = makeStyles({
  chatHeader: {
    display: 'flex',
    alignItems: 'center',
    columnGap: '4px',
    lineHeight: '32px',
    fontSize: '18px',
    fontWeight: 700
  },
  container: {
    display: 'flex',
    flexDirection: 'column',
    gridRowGap: '16px',
    minWidth: '300px',
    backgroundColor: tokens.colorNeutralBackground1
  },
  formButtons: {
    display: 'flex',
    flexDirection: 'row',
    justifyContent: 'flex-end',
    gridColumnGap: '8px'
  }
});

const reduceToFirstNamesList = (participants: AadUserConversationMember[] = [], userId = '') => {
  return participants
    .filter(p => p.userId !== userId)
    .map(p => {
      if (p.displayName?.includes(' ')) {
        return p.displayName.split(' ')[0];
      }
      return p.displayName || p.email || p.id;
    })
    .join(', ');
};

const GroupChatHeader = ({ chat, currentUserId, onRenameChat }: ChatHeaderProps) => {
  const styles = useStyles();
  const [isPopoverOpen, setIsPopoverOpen] = useState(false);
  const handleOpenChange = useCallback((_, data: OnOpenChangeData) => setIsPopoverOpen(data.open || false), []);
  const onCancelClicked = useCallback(() => setIsPopoverOpen(false), []);
  const [chatName, setChatName] = useState(chat?.topic);
  const onChatNameChanged = useCallback(
    (_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, data: InputOnChangeData) => {
      setChatName(data.value);
    },
    []
  );
  const renameChat = useCallback(() => {
    void onRenameChat(chatName || null);
  }, [chatName, onRenameChat]);
  const chatTitle = chat?.topic
    ? chat.topic
    : reduceToFirstNamesList(chat?.members as AadUserConversationMember[], currentUserId);
  return (
    <>
      {chatTitle}
      <Popover trapFocus positioning={'below-start'} open={isPopoverOpen} onOpenChange={handleOpenChange}>
        <PopoverTrigger>
          <Button appearance="transparent" icon={<EditIcon />} aria-label="Name group chat"></Button>
        </PopoverTrigger>
        <PopoverSurface>
          <div className={styles.container}>
            <Field label="Group name">
              <Input placeholder="Enter group name" onChange={onChatNameChanged} value={chatName || ''} />
            </Field>
            <div className={styles.formButtons}>
              <Button appearance="secondary" onClick={onCancelClicked}>
                Cancel
              </Button>
              <Button appearance="primary" disabled={chatName === chat?.topic} onClick={renameChat}>
                Send
              </Button>
            </div>
          </div>
        </PopoverSurface>
      </Popover>
    </>
  );
};

const OneToOneChatHeader = ({ chat, currentUserId }: ChatHeaderProps) => {
  const id = getOtherParticipantUserId(chat, currentUserId);
  return id ? (
    <Person
      userId={id}
      view={ViewType.oneline}
      avatarSize="small"
      personCardInteraction={PersonCardInteraction.hover}
      showPresence={true}
    />
  ) : null;
};

const getOtherParticipantUserId = (chat?: Chat, currentUserId = '') =>
  (chat?.members as AadUserConversationMember[])?.find(m => m.userId !== currentUserId)?.userId;

/**
 * Shows the "name" of the chat.
 * For group chats where the topic is set it will show the topic.
 * For group chats with no topic it will show the first names of other participants comma separated.
 * For 1:1 chats it will show the first name of the other participant.
 */
const ChatHeader = memo(({ chat, currentUserId, onRenameChat }: ChatHeaderProps) => {
  const styles = useStyles();
  return (
    <div className={styles.chatHeader}>
      {chat?.chatType === 'group' ? (
        <GroupChatHeader chat={chat} currentUserId={currentUserId} onRenameChat={onRenameChat} />
      ) : (
        <OneToOneChatHeader chat={chat} currentUserId={currentUserId} onRenameChat={onRenameChat} />
      )}
    </div>
  );
});

export default ChatHeader;
