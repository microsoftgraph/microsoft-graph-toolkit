import React, { useCallback, useState } from 'react';
import {
  makeStyles,
  mergeClasses,
  shorthands,
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle
} from '@fluentui/react-components';
import { List, ListItem } from '@fluentui/react-migration-v0-v9';
import { Person } from '@microsoft/mgt-react';
import { AadUserConversationMember } from '@microsoft/microsoft-graph-types';
import { Dismiss20Filled, bundleIcon, iconFilledClassName, iconRegularClassName } from '@fluentui/react-icons';

interface ListChatMembersProps {
  currentUserId: string;
  members: AadUserConversationMember[];
  removeChatMember: (membershipId: string) => Promise<void>;
  closeParentPopover: () => void;
}

const RemovePerson = bundleIcon(Dismiss20Filled, () => <></>);

const useStyles = makeStyles({
  iconPlaceholder: {
    display: 'flex',
    width: '24px'
  },
  listItem: {
    listStyleType: 'none',
    width: '100%',
    ':focus-visible': {
      [`& .${iconFilledClassName}`]: {
        display: 'inline',
        color: 'var(--colorNeutralForeground2BrandHover)'
      },
      [`& .${iconRegularClassName}`]: {
        display: 'none'
      }
    }
  },
  memberList: {
    fontWeight: 800,
    gridGap: '8px',
    ...shorthands.marginBlock('0'),
    ...shorthands.padding('0')
  },
  personRow: {
    display: 'flex',
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    ...shorthands.paddingBlock('5px'),
    ...shorthands.paddingInline('16px'),
    cursor: 'pointer',
    ':hover': {
      backgroundColor: 'var(--colorSubtleBackgroundHover)',
      [`& .${iconFilledClassName}`]: {
        display: 'inline',
        color: 'var(--colorNeutralForeground2BrandHover)'
      },
      [`& .${iconRegularClassName}`]: {
        display: 'none'
      }
    }
  },
  fullWidth: {
    width: '100%'
  }
});

const ListChatMembers = ({ members, currentUserId, removeChatMember, closeParentPopover }: ListChatMembersProps) => {
  const styles = useStyles();
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
        {members.map(member => {
          const isCurrentUser = currentUserId === member.userId;
          return member?.userId ? (
            <ListItem
              key={member.userId}
              className={styles.listItem}
              onClick={() => {
                if (isCurrentUser) {
                  closeDialog();
                } else {
                  openRemoveDialog(member);
                }
              }}
            >
              <div className={mergeClasses(styles.personRow, styles.fullWidth)}>
                <Person className={styles.fullWidth} tabIndex={-1} userId={member.userId} view="oneline" showPresence />
                <span className={styles.iconPlaceholder}>{!isCurrentUser && <RemovePerson />}</span>
              </div>
            </ListItem>
          ) : null;
        })}
      </List>
    </>
  );
};

export { ListChatMembers };
