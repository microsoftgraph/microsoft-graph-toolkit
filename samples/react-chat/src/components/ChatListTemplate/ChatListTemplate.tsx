import React from 'react';
import { MgtTemplateProps } from '@microsoft/mgt-react';
import { Chat } from '@microsoft/microsoft-graph-types';
import ChatItem, { ChatInteractionProps } from '../ChatItem/ChatItem';

const badIdChat: Chat = {
  // "id": "19:2067b733-8159-4f6e-acb4-960d1307afc2_ad45acdf-5bd2-4f56-9155-2d77bfd9dfe2@unq.gbl.spaces",
  id: '',
  topic: 'No chat ID chat',
  createdDateTime: '2024-01-15T12:01:10.176Z',
  lastUpdatedDateTime: '2024-01-15T12:01:10.176Z',
  chatType: 'oneOnOne',
  webUrl:
    'https://teams.microsoft.com/l/chat/19%3A2067b733-8159-4f6e-acb4-960d1307afc2_ad45acdf-5bd2-4f56-9155-2d77bfd9dfe2%40unq.gbl.spaces/0?tenantId=420b9de2-e599-4222-87b4-4e5bc86ef9e4',
  tenantId: '420b9de2-e599-4222-87b4-4e5bc86ef9e4',
  onlineMeetingInfo: null,
  viewpoint: {
    isHidden: false,
    lastMessageReadDateTime: '2024-01-15T12:01:12.254Z'
  },
  'members@odata.context':
    "https://graph.microsoft.com/v1.0/$metadata#users('2067b733-8159-4f6e-acb4-960d1307afc2')/chats('19%3A2067b733-8159-4f6e-acb4-960d1307afc2_ad45acdf-5bd2-4f56-9155-2d77bfd9dfe2%40unq.gbl.spaces')/members",
  // members: [
  //   {
  //     '@odata.type': '#microsoft.graph.aadUserConversationMember',
  //     id: 'MCMjMCMjNDIwYjlkZTItZTU5OS00MjIyLTg3YjQtNGU1YmM4NmVmOWU0IyMxOToyMDY3YjczMy04MTU5LTRmNmUtYWNiNC05NjBkMTMwN2FmYzJfYWQ0NWFjZGYtNWJkMi00ZjU2LTkxNTUtMmQ3N2JmZDlkZmUyQHVucS5nYmwuc3BhY2VzIyMyMDY3YjczMy04MTU5LTRmNmUtYWNiNC05NjBkMTMwN2FmYzI=',
  //     roles: ['owner'],
  //     displayName: 'Isaiah Langer',
  //     visibleHistoryStartDateTime: '0001-01-01T00:00:00Z',
  //     userId: '2067b733-8159-4f6e-acb4-960d1307afc2',
  //     email: 'IsaiahL@M365x64818216.OnMicrosoft.com',
  //     tenantId: '420b9de2-e599-4222-87b4-4e5bc86ef9e4'
  //   },
  //   {
  //     '@odata.type': '#microsoft.graph.aadUserConversationMember',
  //     id: 'MCMjMCMjNDIwYjlkZTItZTU5OS00MjIyLTg3YjQtNGU1YmM4NmVmOWU0IyMxOToyMDY3YjczMy04MTU5LTRmNmUtYWNiNC05NjBkMTMwN2FmYzJfYWQ0NWFjZGYtNWJkMi00ZjU2LTkxNTUtMmQ3N2JmZDlkZmUyQHVucS5nYmwuc3BhY2VzIyNhZDQ1YWNkZi01YmQyLTRmNTYtOTE1NS0yZDc3YmZkOWRmZTI=',
  //     roles: ['owner'],
  //     displayName: 'Lidia Holloway',
  //     visibleHistoryStartDateTime: '0001-01-01T00:00:00Z',
  //     userId: 'ad45acdf-5bd2-4f56-9155-2d77bfd9dfe2',
  //     email: 'LidiaH@M365x64818216.OnMicrosoft.com',
  //     tenantId: '420b9de2-e599-4222-87b4-4e5bc86ef9e4'
  //   }
  // ]
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
      <ChatItem key={badIdChat.id} chat={badIdChat} onSelected={props.onSelected} />
    </ul>
  );
};

export default ChatListTemplate;
