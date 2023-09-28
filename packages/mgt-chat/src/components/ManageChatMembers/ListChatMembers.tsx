import React, { useCallback, useState } from 'react';
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle
} from '@fluentui/react-components';
import { List, ListItem } from '@fluentui/react-northstar';
import { Person, PersonViewType } from '@microsoft/mgt-react';
import { AadUserConversationMember } from '@microsoft/microsoft-graph-types';
import { styles } from './manage-chat-members.styles';
import { Dismiss24Regular, bundleIcon } from '@fluentui/react-icons';

interface ListChatMembersProps {
  currentUserId: string;
  members: AadUserConversationMember[];
  removeChatMember: (membershipId: string) => Promise<void>;
  closeParentPopover: () => void;
}

const RemovePerson = bundleIcon(Dismiss24Regular, () => <></>);

const ListChatMembers = ({ members, currentUserId, removeChatMember, closeParentPopover }: ListChatMembersProps) => {
  const [removeDialogOpen, setRemoveDialogOpen] = useState(false);
  const [removeUser, setRemoveUser] = useState<AadUserConversationMember | undefined>(undefined);
  const openRemoveDialog = useCallback((user: AadUserConversationMember) => {
    if (!user) return;
    setRemoveUser(user);
    setRemoveDialogOpen(true);
  }, []);

  const closeDialog = useCallback(() => {
    setRemoveUser(undefined);
    setRemoveDialogOpen(false);
    closeParentPopover();
  }, [closeParentPopover]);

  const removeMember = useCallback(() => {
    if (!removeUser?.id) {
      closeDialog();
      return;
    }

    void removeChatMember(removeUser.id).then(closeDialog);
  }, [removeUser, removeChatMember, closeDialog]);
  return (
    <>
      <Dialog open={removeDialogOpen} onOpenChange={(_, data) => setRemoveDialogOpen(data.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Remove {removeUser?.displayName || ''} from the conversation?</DialogTitle>
            <DialogContent>They&apos;ll still have access to the chat history.</DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={closeDialog}>
                Close
              </Button>
              <Button appearance="primary" onClick={removeMember}>
                Remove
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
      <List navigable className={styles.memberList}>
        {members.map((member, index) => {
          const isCurrentUser = currentUserId === member.userId;
          return member?.userId ? (
            <ListItem
              key={member.userId}
              className={styles.listItem}
              index={index}
              content={
                <Button
                  appearance="subtle"
                  icon={isCurrentUser ? null : <RemovePerson />}
                  iconPosition="after"
                  className={styles.fullWidth}
                >
                  <Person
                    className={styles.fullWidth}
                    tabIndex={-1}
                    userId={member.userId}
                    view={PersonViewType.oneline}
                    showPresence
                  />
                </Button>
              }
              onClick={() => {
                if (isCurrentUser) return;
                openRemoveDialog(member);
              }}
            />
          ) : null;
        })}
      </List>
    </>
  );
};

export { ListChatMembers };
