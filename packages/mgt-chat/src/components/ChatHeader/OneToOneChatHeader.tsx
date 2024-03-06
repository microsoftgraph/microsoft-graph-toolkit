import React from 'react';
import { AadUserConversationMember, Chat } from '@microsoft/microsoft-graph-types';
import { Person } from '@microsoft/mgt-react';
import { ChatHeaderProps } from './ChatTitle';
import { makeStyles } from '@fluentui/react-components';
import { useBotInfo } from '../../statefulClient/useBotInfo';
import { BotAvatar } from '../ChatAvatar/BotAvatar';

const getOtherParticipantUserId = (chat?: Chat, currentUserId = '') =>
  (chat?.members as AadUserConversationMember[])?.find(m => m.userId !== currentUserId)?.userId;

const useStyles = makeStyles({
  person: {
    '--person-avatar-size': '32px',
    '--person-alignment': 'center'
  }
});
export const OneToOneChatHeader = ({ chat, currentUserId }: ChatHeaderProps) => {
  const botInfo = useBotInfo();
  const styles = useStyles();
  const id = getOtherParticipantUserId(chat, currentUserId);
  if (!chat?.id) return null;
  if (id) {
    return (
      <Person
        className={styles.person}
        userId={id}
        view="oneline"
        avatarSize="small"
        personCardInteraction="hover"
        showPresence={true}
      />
    );
  } else if (botInfo?.chatBots.has(chat.id)) {
    return (
      <>
        {Array.from(botInfo.chatBots.get(chat.id)!).map(bot =>
          bot.teamsAppDefinition?.bot?.id ? (
            <BotAvatar
              key={bot.teamsAppDefinition.bot.id}
              chatId={chat.id!}
              avatarId={bot.teamsAppDefinition.bot.id}
              view="oneline"
            />
          ) : null
        )}
      </>
    );
  }
  return null;
};
