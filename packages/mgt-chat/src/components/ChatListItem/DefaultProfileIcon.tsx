import React from 'react';
import { makeStyles, shorthands } from '@fluentui/react-components';
import { ChatListItemIcon } from '../ChatListItemIcon/ChatListItemIcon';
import { MgtTemplateProps } from '@microsoft/mgt-react';

const useStyles = makeStyles({
  defaultProfileImage: {
    ...shorthands.borderRadius('50%'),
    objectFit: 'cover',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center'
  }
});

export const DefaultProfileIcon = (_: MgtTemplateProps) => {
  const styles = useStyles();
  const oneOnOneProfilePicture = <ChatListItemIcon chatType="oneOnOne" />;
  return <div className={styles.defaultProfileImage}>{oneOnOneProfilePicture}</div>;
};
