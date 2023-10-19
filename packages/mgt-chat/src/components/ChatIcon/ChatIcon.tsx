import React from 'react';
import { Chat } from '@microsoft/microsoft-graph-types';
import { Calendar16Regular, PeopleTeam16Regular, bundleIcon } from '@fluentui/react-icons';
import { error } from '@microsoft/mgt-element';
import { Circle } from '../Circle/Circle';

const MeetingIcon = bundleIcon(Calendar16Regular, Calendar16Regular);
const GroupIcon = bundleIcon(PeopleTeam16Regular, PeopleTeam16Regular);
export const ChatIcon = ({ chatType }: Chat): JSX.Element | null => {
  if (!chatType) return null;

  const iconColor = 'var(--colorBrandForeground2)';

  switch (chatType) {
    case 'group':
      return (
        <Circle>
          <GroupIcon color={iconColor} />
        </Circle>
      );
    case 'meeting':
      return (
        <Circle>
          <MeetingIcon color={iconColor} />
        </Circle>
      );
    default:
      error(`Attempted to render an icon for chat of type: ${chatType}`);
      return null;
  }
};
