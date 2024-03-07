import React from 'react';
import { Chat } from '@microsoft/microsoft-graph-types';
import { makeStyles, shorthands } from '@fluentui/react-components';
import { CalendarLtr24Regular, PeopleTeam24Regular, Person24Regular, bundleIcon } from '@fluentui/react-icons';
import { error } from '@microsoft/mgt-element';
import { Circle } from '../Circle/Circle';
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

const GroupIcon = bundleIcon(PeopleTeam24Regular, PeopleTeam24Regular);
const PersonIcon = bundleIcon(Person24Regular, Person24Regular);
const MeetingIcon = bundleIcon(CalendarLtr24Regular, CalendarLtr24Regular);

export const ChatListItemIcon = ({ chatType }: Chat & MgtTemplateProps): JSX.Element | null => {
  const styles = useStyles();
  if (!chatType) return null;

  const iconColor = 'var(--colorBrandForeground2)';

  switch (chatType) {
    case 'meeting':
      return (
        <Circle>
          <MeetingIcon color={iconColor} />
        </Circle>
      );
    case 'oneOnOne':
      return (
        <div className={styles.defaultProfileImage}>
          <Circle>
            <PersonIcon color={iconColor} />
          </Circle>
        </div>
      );
    case 'group':
      return (
        <Circle>
          <GroupIcon color={iconColor} />
        </Circle>
      );
    default:
      error(`Attempted to render an icon for chat of type: ${chatType}`);
      return null;
  }
};
