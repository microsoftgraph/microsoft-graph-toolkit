import React, { useCallback } from 'react';
import { MgtTemplateProps } from '@microsoft/mgt-react';
import { makeStyles, shorthands, Button } from '@fluentui/react-components';
import { EllipsisMenu, IChatListMenuItemsProps } from './EllipsisMenu';
import { ChatListButtonItem } from './ChatListButtonItem';
import { Circle } from '../Circle/Circle';
import { IChatListActions } from './IChatListActions';

const useStyles = makeStyles({
  headerContainer: {
    display: 'flex',
    width: '100%',
    flexDirection: 'row',
    rowGap: 0,
    zIndex: 3,
    paddingTop: '8px',
    paddingLeft: '10px',
    paddingRight: '10px',
    paddingBlockEnd: '8px'
  },
  controlsContainer: {
    display: 'flex',
    flexDirection: 'row',
    justifyContent: 'space-between', // This distributes the space evenly
    alignItems: 'center'
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
      actions: IChatListActions;
    }
) => {
  const classes = useStyles();

  const buttonItems: ChatListButtonItem[] = props.buttonItems === undefined ? [] : props.buttonItems;
  const clickButton = useCallback((buttonItem: ChatListButtonItem) => {
    buttonItem.onClick(props.actions);
  }, []);

  return (
    <>
      {buttonItems.length === 0 && (props.menuItems === undefined || props.menuItems?.length === 0) ? (
        <></>
      ) : (
        <div className={classes.headerContainer}>
          <div className={classes.controlsContainer}>
            <div>
              {buttonItems.map((buttonItem, index) => (
                <Button key={index} className={classes.button} onClick={() => clickButton(buttonItem)}>
                  <div className={classes.buttonIcon}>
                    <Circle>{buttonItem.renderIcon()}</Circle>
                  </div>
                </Button>
              ))}
            </div>
            <div>
              <EllipsisMenu actions={props.actions} menuItems={props.menuItems} />
            </div>
          </div>
        </div>
      )}
    </>
  );
};
