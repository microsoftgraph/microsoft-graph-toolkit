import React from 'react';
import { MgtTemplateProps, Login } from '@microsoft/mgt-react';
import { makeStyles, shorthands, Button } from '@fluentui/react-components';
import { EllipsisMenu, IChatListMenuItemsProps } from './EllipsisMenu';
import { ChatListButtonItem } from './ChatListButtonItem';
import { Circle } from '../Circle/Circle';

const useStyles = makeStyles({
  headerContainer: {
    display: 'flex',
    flexDirection: 'row',
    justifyContent: 'space-between', // This distributes the space evenly
    alignItems: 'center',
    width: '100%'
  },
  buttonIcon: {
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

// this is a stub to move the logic here that should end up here.
export const ChatListHeader = (
  props: MgtTemplateProps &
    IChatListMenuItemsProps & {
      buttonItems?: ChatListButtonItem[];
    }
) => {
  const classes = useStyles();

  const buttonItems: ChatListButtonItem[] = props.buttonItems === undefined ? [] : props.buttonItems;
  return (
    <div className={classes.headerContainer}>
      <div>
        {buttonItems.map((buttonItem, index) => (
          <Button key={index} className={classes.button} onClick={buttonItem.onClick}>
            <div className={classes.buttonIcon}>
              <Circle>{buttonItem.renderIcon()}</Circle>
            </div>
          </Button>
        ))}
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
