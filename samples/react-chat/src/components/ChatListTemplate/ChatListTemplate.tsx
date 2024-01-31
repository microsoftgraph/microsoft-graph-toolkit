import React from 'react';
import { MgtTemplateProps } from '@microsoft/mgt-react';
import { Chat } from '@microsoft/microsoft-graph-types';
import ChatItem, { ChatInteractionProps } from '../ChatItem/ChatItem';

const noChatId: Chat = {
  // "id": "19:2067b733-8159-4f6e-acb4-960d1307afc2_ad45acdf-5bd2-4f56-9155-2d77bfd9dfe2@unq.gbl.spaces",
  id: '',
  topic: 'No chat ID chat',
  createdDateTime: '2024-01-15T12:01:10.176Z',
  lastUpdatedDateTime: '2024-01-15T12:01:10.176Z',
  chatType: 'oneOnOne',
  webUrl: undefined,
  tenantId: '420b9de2-e599-4222-87b4-4e5bc86ef9e4',
  onlineMeetingInfo: null,
  viewpoint: {
    isHidden: false,
    lastMessageReadDateTime: '2024-01-15T12:01:12.254Z'
  },
  'members@odata.context': undefined,
  members: undefined
} as Chat;

const badChatId: Chat = {
  id: '19:2067b733-8159-4f6e-xxxx-xxxx',
  topic: 'Wrong chat ID chat',
  createdDateTime: '2024-01-15T12:01:10.176Z',
  lastUpdatedDateTime: '2024-01-15T12:01:10.176Z',
  chatType: 'oneOnOne',
  webUrl: undefined,
  tenantId: '420b9de2-e599-4222-87b4-4e5bc86ef9e4',
  onlineMeetingInfo: null,
  viewpoint: {
    isHidden: false,
    lastMessageReadDateTime: '2024-01-15T12:01:12.254Z'
  },
  'members@odata.context': undefined,
  members: undefined
} as Chat;

const ChatListTemplate = (props: MgtTemplateProps & ChatInteractionProps) => {
  const { value } = props.dataContext as { value: Chat[] };
  const chats: Chat[] = value;
  // Select a default chat to display
  // props.onSelected(chats[0]);
  return (
    <ul>
      {chats.map(c => (
        <ChatItem key={c.id} chat={c} onSelected={props.onSelected} />
      ))}
      <ChatItem key={noChatId.id} chat={noChatId} onSelected={props.onSelected} />
      <ChatItem key={badChatId.id} chat={badChatId} onSelected={props.onSelected} />
    </ul>
  );
};

export default ChatListTemplate;
