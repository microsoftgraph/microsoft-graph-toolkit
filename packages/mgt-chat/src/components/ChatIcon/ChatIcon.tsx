import React from 'react';
import { Chat } from '@microsoft/microsoft-graph-types';
import {
  Calendar16Regular,
  PeopleTeam16Regular,
  bundleIcon,
  Person16Regular,
  Person16Filled
} from '@fluentui/react-icons';
import { Circle } from '../Circle/Circle';

const MeetingIcon = bundleIcon(Calendar16Regular, Calendar16Regular);
const GroupIcon = bundleIcon(PeopleTeam16Regular, PeopleTeam16Regular);
const PersonIcon = bundleIcon(Person16Filled, Person16Regular);

export const ChatIcon = ({ chatType }: Chat): JSX.Element | null => {
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
      return (
        <Circle>
          <PersonIcon color={iconColor} />
        </Circle>
      );
  }
};
