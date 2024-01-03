import React from 'react';
import { MgtTemplateProps, Login } from '@microsoft/mgt-react';
import { ChatAdd24Filled, ChatAdd24Regular, bundleIcon } from '@fluentui/react-icons';
import { makeStyles } from '@fluentui/react-components';
import { Circle } from '../Circle/Circle';
import { EllipsisMenu } from './EllipsisMenu';
const ChatAddIconBundle = bundleIcon(ChatAdd24Filled, ChatAdd24Regular);

const useStyles = makeStyles({
  headerContainer: {
    display: 'flex',
    flexDirection: 'row',
    justifyContent: 'space-between', // This distributes the space evenly
    alignItems: 'center',
    width: '100%'
  }
});

export const ChatAddIcon = (): JSX.Element => {
  const iconColor = 'var(--colorBrandForeground2)';
  return (
    <Circle>
      <ChatAddIconBundle color={iconColor} />
    </Circle>
  );
};

// this is a stub to move the logic here that should end up here.
export const ChatListHeader = (props: MgtTemplateProps) => {
  const classes = useStyles();

  return (
    <div className={classes.headerContainer}>
      <div>
        <ChatAddIcon />
      </div>
      <div>
        <Login showPresence={true} loginView="avatar" />
      </div>
      <div>
        <EllipsisMenu
          items={[
            {
              displayText: 'Mark all as read'
            }
          ]}
        />
      </div>
    </div>
  );
};

export default ChatListHeader;
