import { Person, ViewType } from '@microsoft/mgt-react';
import React, { FC, useEffect } from 'react';
import { useBotInfo } from '../../statefulClient/useBotInfo';
import { ChatAvatarProps } from './ChatAvatar';

type BotAvatarProps = ChatAvatarProps & {
  view?: ViewType;
};

export const BotAvatar: FC<BotAvatarProps> = ({ chatId, avatarId, view = 'image' }) => {
  const botInfo = useBotInfo();

  useEffect(() => {
    if (chatId && avatarId && !botInfo?.botInfo.has(avatarId)) {
      void botInfo?.loadBotInfo(chatId, avatarId);
    }
  }, [chatId, avatarId, botInfo]);

  return (
    <div>
      <Person
        personDetails={{
          id: avatarId,
          displayName: botInfo?.botInfo.has(avatarId)
            ? botInfo.botInfo.get(avatarId)?.teamsAppDefinition?.displayName
            : '',
          personImage: botInfo?.botIcons.has(avatarId) ? botInfo.botIcons.get(avatarId) : ''
        }}
        avatarSize="small"
        personCardInteraction="none"
        view={view}
        showPresence={false}
      />
    </div>
  );
};
