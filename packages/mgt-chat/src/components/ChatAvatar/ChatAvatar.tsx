import { Person } from '@microsoft/mgt-react';
import React, { FC, memo, useContext, useEffect, useState } from 'react';
import { botPrefix } from '../../statefulClient/buildBotId';
import { BotInfoContext } from '../Context/BotInfoContext';

interface ChatAvatarProps {
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

const BotAvatar: FC<ChatAvatarProps> = ({ chatId, avatarId }) => {
  const botInfoClient = useContext(BotInfoContext);
  const [botInfo, setBotInfo] = useState(() => botInfoClient?.getState());
  useEffect(() => {
    if (botInfoClient) {
      botInfoClient.onStateChange(setBotInfo);
      return () => botInfoClient.offStateChange(setBotInfo);
    }
  }, [botInfoClient]);

  useEffect(() => {
    if (chatId && avatarId && !botInfo?.botInfo.has(avatarId)) {
      void botInfo?.loadBotInfo(chatId, avatarId);
    }
  }, [chatId, avatarId, botInfo]);

  return (
    <div>
      {/* {botInfo?.botInfo.has(avatarId) ? botInfo.botInfo.get(avatarId)?.teamsAppDefinition?.displayName || '' : null} */}
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
        showPresence={false}
      />
    </div>
  );
};

const AvatarSwitcher: FC<ChatAvatarProps> = ({ chatId, avatarId }) =>
  avatarId.startsWith(botPrefix) ? (
    <BotAvatar chatId={chatId} avatarId={avatarId.replace(botPrefix, '')} />
  ) : (
    <Person userId={avatarId} avatarSize="small" personCardInteraction="hover" showPresence={true} />
  );

export const ChatAvatar = memo(AvatarSwitcher);
