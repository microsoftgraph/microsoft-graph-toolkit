import { Persona, makeStyles, mergeClasses, shorthands } from '@fluentui/react-components';
import { MgtTemplateProps, Person, ViewType } from '@microsoft/mgt-react';
import { Chat, AadUserConversationMember } from '@microsoft/microsoft-graph-types';
import React, { useCallback, useState } from 'react';
import { PeopleTeam16Regular, Calendar16Regular } from '@fluentui/react-icons';

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
    userSelect: 'none',
    '--person-line1-font-weight': 'var(--fontWeightRegular)',
    '--person-avatar-size-small': '40px',
    '--person-avatar-size': '40px',
    '--person-line2-font-size': 'var(--fontSizeBase200)',
    '--person-line2-text-color': 'var(--colorNeutralForeground4)',
    '& .fui-Persona__primaryText': {
      fontSize: 'var(--fontSizeBase300)',
      fontWeight: 'var(--fontWeightRegular)',
    },
    '& .fui-Persona__secondaryText': {
      whiteSpace: 'nowrap',
      textOverflow: 'ellipsis',
      width: '200px',
      display: 'inline-block',
      fontSize: 'var(--fontSizeBase200);',
      color: 'var(--colorNeutralForeground4)',
      ...shorthands.overflow('hidden')
    }
  },
  group: {
    paddingTop: '5px',
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
  userId?: string;
}

interface MessagePreviewProps {
  messagePreview?: string;
}

const getMessagePreview = (chat: Chat, userId?: string): string | undefined => {

  let preview = "";
  if (chat?.lastMessagePreview?.from?.user && chat?.lastMessagePreview?.from?.user.id !== userId) {
    preview += chat?.lastMessagePreview?.from?.user.displayName?.split(' ')[0]!;
    preview += ": ";
  } else {
    preview += "You: ";
  }

  if (chat?.lastMessagePreview?.body?.contentType === 'text') {
    preview += chat?.lastMessagePreview?.body?.content;
  } else if (chat?.lastMessagePreview?.body?.contentType === 'html') {
    const html = chat?.lastMessagePreview?.body?.content!;
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');

    const images = doc.querySelectorAll('img');
    const card = doc.querySelector('attachment');
    const systemCard = doc.querySelector('span[itemId]');
    const systemEventMessage = doc.querySelector('systemEventMessage');

    if (systemEventMessage || systemCard) return undefined;

    preview += doc.body.textContent || doc.body.innerText || "";

    if (images) {
      if (Array.from(images.values()).find(i => i.src.includes('.gif'))) {
        preview += " 📷 GIF";
      } else {
        preview += " Sent an image";
      }
    }

    if (card) {
      preview = "Sent a card";
    }
  }

  return preview?.trim().replace(/[\n\t\r]/g, "").replace(/\s+/g, ' ');
};

const ChatItem = React.memo(({ chat, isSelected, onSelected, userId }: ChatItemProps & ChatInteractionProps) => {
  const styles = useStyles();
  const [messagePreview, setMessagePreview] = useState<string | undefined>("");

  const getOtherParticipantId = useCallback(
    (chat: Chat) => {
      const member = chat.members?.find(m => (m as AadUserConversationMember).userId !== userId);

      if (member) {
        return (member as AadUserConversationMember).userId as string;
      } else if (chat.members?.length === 1 && (chat.members[0] as AadUserConversationMember).userId === userId) {
        return userId;
      }

      return undefined;
    },
    [userId]
  );

  const getGroupTitle = useCallback((chat: Chat) => {
    const lf = new Intl.ListFormat('en');
    let groupMembers: string[] = [];

    if (chat.topic) {
      return chat.topic;
    } else {
      chat.members?.filter(member => member["userId"] !== userId)
        ?.forEach(member => {
          groupMembers.push(member.displayName?.split(' ')[0]!);
        });

      return lf.format(groupMembers);
    }
  }, [userId]);

  React.useEffect(() => {
    setMessagePreview(getMessagePreview(chat, userId));
  }, [chat, userId]);

  return (
    <>
      {userId && (
        <div className={mergeClasses(styles.chat, `${isSelected && styles.active}`)}>
          {chat.chatType === 'oneOnOne' && (
            <Person
              userId={getOtherParticipantId(chat)}
              view={messagePreview ? ViewType.twolines : ViewType.oneline}
              avatarSize="auto"
              showPresence={true}
              onClick={() => onSelected(chat)}
              className={styles.person}
            >
              {messagePreview && (
                <MessagePreview template="line2" chat={chat} userId={userId} messagePreview={messagePreview} />
              )}
            </Person>
          )}
          {chat.chatType === 'group' && (
            <div onClick={() => onSelected(chat)}>
              <Persona
                textAlignment="center"
                size="extra-large"
                name={getGroupTitle(chat)}
                secondaryText={messagePreview}
                avatar={{ icon: <PeopleTeam16Regular />, initials: null }}
                className={mergeClasses(styles.person, styles.group)}
              />
              <span></span>
            </div>
          )}
          {chat.chatType === 'meeting' && (
            <div onClick={() => onSelected(chat)}>
              <Persona
                textAlignment="center"
                size="extra-large"
                className={mergeClasses(styles.person, styles.group)}
                avatar={{ icon: <Calendar16Regular />, initials: null }}
                name={getGroupTitle(chat)}
                secondaryText={messagePreview}
              />
            </div>
          )}
        </div>
      )}
    </>
  );
});

const MessagePreview = (props: MgtTemplateProps & ChatItemProps & MessagePreviewProps) => {
  const styles = useStyles();

  return (
    <>
      {props.messagePreview && (
        <span className={styles.messagePreview}>{props.messagePreview}</span>
      )}
    </>
  );
};

export default React.memo(ChatItem);
