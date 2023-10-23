import React, { useCallback, useState } from 'react';
import { AadUserConversationMember } from '@microsoft/microsoft-graph-types';
import { Edit16Filled, Edit16Regular, bundleIcon } from '@fluentui/react-icons';
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
import type { PopoverProps } from '@fluentui/react-components';
import { ChatIcon } from '../ChatIcon/ChatIcon';
import { ChatHeaderProps } from './ChatTitle';

const EditIcon = bundleIcon(Edit16Filled, Edit16Regular);
const useStyles = makeStyles({
  chatTitle: {
    marginInlineStart: '10px'
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
export const GroupChatHeader = ({ chat, currentUserId, onRenameChat }: ChatHeaderProps) => {
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
  const trapFocus = true;
  const popoverProps: Partial<PopoverProps> = {
    trapFocus,
    inertTrapFocus: trapFocus,
    inline: true,
    positioning: 'below-start',
    open: isPopoverOpen,
    onOpenChange: handleOpenChange
  };
  return (
    <>
      <ChatIcon chatType={chat?.chatType} />
      <span className={styles.chatTitle}>{chatTitle}</span>
      <Popover {...popoverProps}>
        <PopoverTrigger>
          <Button appearance="transparent" icon={<EditIcon />} aria-label="Name group chat"></Button>
        </PopoverTrigger>
        <PopoverSurface>
          <div className={styles.container}>
            <Field label="Group name">
              <Input placeholder="Enter group name" onChange={onChatNameChanged} value={chatName || ''} />
            </Field>
            <div className={styles.formButtons}>
              <Button appearance="secondary" onClick={onCancelClicked} aria-label="cancel">
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
