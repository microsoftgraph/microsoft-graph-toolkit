import React from 'react';
import { MgtTemplateProps, Login } from '@microsoft/mgt-react';
import { ChatAdd24Filled, ChatAdd24Regular, bundleIcon } from '@fluentui/react-icons';
import { makeStyles, shorthands, Button } from '@fluentui/react-components';
import { Circle } from '../Circle/Circle';
import { IChatListMenuItemsProps, EllipsisMenu } from './EllipsisMenu';
const ChatAddIconBundle = bundleIcon(ChatAdd24Filled, ChatAdd24Regular);

const useStyles = makeStyles({
  headerContainer: {
    display: 'flex',
    flexDirection: 'row',
    justifyContent: 'space-between', // This distributes the space evenly
    alignItems: 'center',
    width: '100%'
  },
  chatAddIcon: {
    flexGrow: 0,
    flexShrink: 0,
    flexBasis: '35px',
    ...shorthands.borderRadius('50%'), // This will make it round
    marginRight: '10px',
    objectFit: 'cover', // This ensures the image covers the area without stretching
    display: 'flex',
    alignItems: 'center', // This will vertically center the image
    justifyContent: 'center' // This will horizontally center the image
  },
  button: {
    ...shorthands.border('none'),
    ...shorthands.padding(0), // Remove padding
    minWidth: 0, // Allow the button to shrink to the size of its content
    '&:hover': {
      backgroundColor: 'transparent'
    }
  }
});

export const ChatAddIcon = (): JSX.Element => {
  const classes = useStyles();
  const iconColor = 'var(--colorBrandForeground2)';
  return (
    <div className={classes.chatAddIcon}>
      <Circle>
        <ChatAddIconBundle color={iconColor} />
      </Circle>
    </div>
  );
};

// this is a stub to move the logic here that should end up here.
export const ChatListHeader = (props: MgtTemplateProps & IChatListMenuItemsProps) => {
  const classes = useStyles();

  return (
    <div className={classes.headerContainer}>
      <div>
        <Button className={classes.button}>
          <ChatAddIcon />
        </Button>
      </div>
      <div>
        <Login showPresence={true} loginView="avatar" />
      </div>
      <div>
        <EllipsisMenu menuItems={props.menuItems} />
      </div>
    </div>
  );
};

export default ChatListHeader;
