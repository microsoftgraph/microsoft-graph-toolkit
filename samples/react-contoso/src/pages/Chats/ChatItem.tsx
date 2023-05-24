import { Persona, makeStyles, mergeClasses, shorthands } from '@fluentui/react-components';
import { Providers } from '@microsoft/mgt-element';
import { MgtTemplateProps, Person, ViewType } from '@microsoft/mgt-react';
import { Chat, AadUserConversationMember } from '@microsoft/microsoft-graph-types';
import React, { useCallback, useEffect, useState } from 'react';
import { PeopleCommunityRegular, CalendarMonthRegular } from '@fluentui/react-icons';

const useStyles = makeStyles({
  chat: {
    paddingLeft: '5px',
    paddingRight: '5px',
    display: 'flex',
    alignItems: 'center',
    height: '50px',
    cursor: 'pointer',
    ':hover': {
      backgroundColor: 'var(--colorNeutralBackground1Hover)'
    }
  },
  active: {
    backgroundColor: 'var(--colorNeutralBackground1Selected)'
  },
  person: {
    '--person-avatar-size-small': '40px',
    '& .fui-Persona__primaryText': {
      fontSize: 'var(--fontSizeBase300);'
    },
    '& .fui-Persona__secondaryText': {
      whiteSpace: 'nowrap',
      textOverflow: 'ellipsis',
      width: '200px',
      display: 'inline-block',
      ...shorthands.overflow('hidden')
    }
  },
  messagePreview: {
    whiteSpace: 'nowrap',
    textOverflow: 'ellipsis',
    width: '200px',
    display: 'inline-block',
    ...shorthands.overflow('hidden')
  }
});

export interface ChatInteractionProps {
  onSelected: (selected: Chat) => void;
  selectedChat?: Chat;
}

interface ChatItemProps {
  chat: Chat;
  isSelected?: boolean;
}

const getMessagePreview = (chat: Chat) => {
  return chat?.lastMessagePreview?.body?.contentType === 'text' ? chat?.lastMessagePreview?.body?.content : '...';
};

const ChatItem = ({ chat, isSelected, onSelected }: ChatItemProps & ChatInteractionProps) => {
  const styles = useStyles();
  const [myId, setMyId] = useState<string>();

  useEffect(() => {
    const getMyId = async () => {
      const me = await Providers.me();
      setMyId(me.id);
    };
    if (!myId) {
      void getMyId();
    }
  }, [myId]);

  const getOtherParticipantId = useCallback(
    (chat: Chat) => {
      const member = chat.members?.find(m => (m as AadUserConversationMember).userId !== myId);

      if (member) {
        console.log('member', member);
        return (member as AadUserConversationMember).userId as string;
      } else if (chat.members?.length === 1 && (chat.members[0] as AadUserConversationMember).userId === myId) {
        return myId;
      }

      return undefined;
    },
    [myId]
  );

  const getGroupTitle = useCallback((chat: Chat) => {
    let groupTitle: string | undefined = '';
    if (chat.topic) {
      groupTitle = chat.topic;
    } else {
      groupTitle = chat.members
        ?.map(member => {
          return member.displayName?.split(' ')[0];
        })
        .join(', ');
    }

    return groupTitle;
  }, []);

  return (
    <>
      {myId && (
        <div className={mergeClasses(styles.chat, `${isSelected && styles.active}`)}>
          {chat.chatType === 'oneOnOne' && (
            <Person
              userId={getOtherParticipantId(chat)}
              view={ViewType.twolines}
              avatarSize="auto"
              showPresence={true}
              onClick={() => onSelected(chat)}
              className={styles.person}
            >
              <MessagePreview template="line2" chat={chat} />
            </Person>
          )}
          {chat.chatType === 'group' && (
            <div onClick={() => onSelected(chat)}>
              <Persona
                textAlignment="center"
                size="extra-large"
                name={getGroupTitle(chat)}
                secondaryText={getMessagePreview(chat)}
                avatar={{ icon: <PeopleCommunityRegular />, initials: null }}
                className={styles.person}
              />
              <span></span>
            </div>
          )}
          {chat.chatType === 'meeting' && (
            <div onClick={() => onSelected(chat)}>
              <Persona
                textAlignment="center"
                size="extra-large"
                className={styles.person}
                avatar={{ icon: <CalendarMonthRegular />, initials: null }}
                name={getGroupTitle(chat)}
                secondaryText={getMessagePreview(chat)}
              />
            </div>
          )}
        </div>
      )}
    </>
  );
};

const MessagePreview = (props: MgtTemplateProps & ChatItemProps) => {
  const styles = useStyles();

  return (
    <>
      <span className={styles.messagePreview}>{getMessagePreview(props.chat)}</span>
    </>
  );
};

export default ChatItem;
