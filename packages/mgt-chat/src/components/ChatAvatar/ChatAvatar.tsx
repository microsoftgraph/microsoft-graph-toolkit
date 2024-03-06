import { Person } from '@microsoft/mgt-react';
import React, { FC, memo } from 'react';
import { botPrefix } from '../../statefulClient/buildBotId';
import { BotAvatar } from './BotAvatar';

export interface ChatAvatarProps {
  /**
   * The chat id
   */
  chatId: string;
  /**
   * The id of the entity to get the avatar for
   * for bots this is prefixed with 'botId::'
   *
   */
  avatarId: string;
}

const AvatarSwitcher: FC<ChatAvatarProps> = ({ chatId, avatarId }) =>
  avatarId.startsWith(botPrefix) ? (
    <BotAvatar chatId={chatId} avatarId={avatarId.replace(botPrefix, '')} />
  ) : (
    <Person userId={avatarId} avatarSize="small" personCardInteraction="hover" showPresence={true} />
  );

export const ChatAvatar = memo(AvatarSwitcher);
