import React, { useCallback, useState } from 'react';
import {
  makeStyles,
  shorthands,
  Button,
  Popover,
  PopoverSurface,
  PopoverTrigger,
  Divider,
  OnOpenChangeData,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  DialogTrigger,
  PopoverProps
} from '@fluentui/react-components';
import {
  bundleIcon,
  PeopleAdd24Regular,
  PeopleAdd24Filled,
  DoorArrowLeft24Filled,
  DoorArrowLeft24Regular
} from '@fluentui/react-icons';
import { AadUserConversationMember } from '@microsoft/microsoft-graph-types';
import { AddChatMembers } from './AddChatMembers';
import { ListChatMembers } from './ListChatMembers';

interface ManageChatMembersProps {
  currentUserId: string;
  members: AadUserConversationMember[];
  removeChatMember: (membershipId: string) => Promise<void>;
  addChatMembers: (userIds: string[], history?: Date) => Promise<void>;
}

const AddPeople = bundleIcon(PeopleAdd24Filled, PeopleAdd24Regular);
const Leave = bundleIcon(DoorArrowLeft24Filled, DoorArrowLeft24Regular);

const useStyles = makeStyles({
  button: {
    justifyContent: 'flex-start',
    ...shorthands.paddingInline('16px')
  },
  buttonRow: {
    display: 'flex',
    flexDirection: 'column'
  },
  container: {
    ...shorthands.paddingBlock('4px')
  },
  divider: {
    ...shorthands.marginBlock('4px')
  },
  popover: {
    ...shorthands.padding('0 !important')
  },
  triggerButton: {
    minWidth: 'unset !important',
    width: 'max-content'
  }
});

const ManageChatMembers = ({ currentUserId, members, addChatMembers, removeChatMember }: ManageChatMembersProps) => {
  const styles = useStyles();
  const [isPopoverOpen, setIsPopoverOpen] = useState(false);
  const [showAddMembers, setShowAddMembers] = useState(false);
  const openAddMembers = useCallback(() => {
    setShowAddMembers(true);
  }, [setShowAddMembers]);
  const closeCallout = useCallback(() => {
    setShowAddMembers(false);
    setIsPopoverOpen(false);
  }, [setShowAddMembers]);
  const handleOpenChange = useCallback(
    (_, data: OnOpenChangeData) => setIsPopoverOpen(data.open || false),
    [setIsPopoverOpen]
  );

  const leaveChat = useCallback(() => {
    const me = members.find(member => member.userId === currentUserId);
    if (me?.id) {
      void removeChatMember(me.id).then(closeCallout);
      return;
    }
    closeCallout();
  }, [removeChatMember, members, currentUserId, closeCallout]);
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
    <Popover {...popoverProps}>
      <PopoverTrigger>
        <Button className={styles.triggerButton} appearance="transparent" icon={<AddPeople />}>
          {members.length}
        </Button>
      </PopoverTrigger>
      <PopoverSurface className={styles.popover}>
        {showAddMembers ? (
          <AddChatMembers closeDialog={closeCallout} addChatMembers={addChatMembers} />
        ) : (
          <div className={styles.container}>
            <ListChatMembers
              members={members}
              removeChatMember={removeChatMember}
              currentUserId={currentUserId}
              closeParentPopover={closeCallout}
            />
            <Divider className={styles.divider} />
            <div className={styles.buttonRow}>
              <Button appearance="subtle" icon={<AddPeople />} onClick={openAddMembers} className={styles.button}>
                Add people
              </Button>
              <Dialog>
                <DialogTrigger>
                  <Button appearance="subtle" icon={<Leave />} className={styles.button}>
                    Leave
                  </Button>
                </DialogTrigger>
                <DialogSurface>
                  <DialogBody>
                    <DialogTitle>Leave the conversation?</DialogTitle>
                    <DialogContent>You&apos;ll still have access to the chat history.</DialogContent>
                    <DialogActions>
                      <DialogTrigger disableButtonEnhancement>
                        <Button appearance="secondary" onClick={closeCallout}>
                          Cancel
                        </Button>
                      </DialogTrigger>
                      {members.length > 2 && (
                        <Button appearance="primary" onClick={leaveChat}>
                          Leave
                        </Button>
                      )}
                    </DialogActions>
                  </DialogBody>
                </DialogSurface>
              </Dialog>
            </div>
          </div>
        )}
      </PopoverSurface>
    </Popover>
  );
};

export { ManageChatMembers };
