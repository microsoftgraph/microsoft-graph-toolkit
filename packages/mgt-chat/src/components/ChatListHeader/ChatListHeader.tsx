import React from 'react';
import { MgtTemplateProps, Login } from '@microsoft/mgt-react';
import { ChatAdd24Filled, ChatAdd24Regular, bundleIcon } from '@fluentui/react-icons';
import { makeStyles } from '@fluentui/react-components';
import { Circle } from '../Circle/Circle';
import { EllipsisMenu } from './EllipsisMenu';
const ChatAddIcon = bundleIcon(ChatAdd24Filled, ChatAdd24Regular);

const useStyles = makeStyles({
  headerContainer: {
    display: 'flex',
    justifyContent: 'space-between', // This distributes the space evenly
    alignItems: 'center',
    width: '300px'
  }
});

// this is a stub to move the logic here that should end up here.
export const ChatListHeader = (props: MgtTemplateProps) => {
  const classes = useStyles();
  const iconColor = 'var(--colorBrandForeground2)';

  return (
    <div className={classes.headerContainer}>
      <div>
        <Circle>
          <ChatAddIcon color={iconColor} />
        </Circle>
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
