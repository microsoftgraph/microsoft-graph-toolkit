import {
  Chat,
  ChatMessageInfo,
  ChatMessageFromIdentitySet,
  AadUserConversationMember,
  MembersAddedEventMessageDetail
} from '@microsoft/microsoft-graph-types';

/* section: oneOnOne SAMPLE DATA */

export const SampleFromAndrew: ChatMessageFromIdentitySet = {
  user: {
    id: 'caab0f18-e467-4d88-ae65-c08f769b8e01',
    displayName: 'Andrew'
  }
};

export const SampleFromDexter: ChatMessageFromIdentitySet = {
  user: {
    id: 'a7530e0a-85f0-4f77-a617-2a5ab2a959ec',
    displayName: 'Dexter'
  }
};

export const SampleMessagePreviewFromAndrew: ChatMessageInfo = {
  createdDateTime: '2021-07-06T21:58:36.141Z',
  body: {
    content: 'This is a sample message preview. If you can send the end of the sentence eol',
    contentType: 'text'
  },
  from: SampleFromAndrew
};

export const SampleMessagePreviewFromDexter: ChatMessageInfo = {
  createdDateTime: '2021-07-06T21:58:36.141Z',
  body: {
    content: 'Note to self',
    contentType: 'text'
  },
  from: SampleFromDexter
};

export const SampleMembers: AadUserConversationMember[] = [
  {
    userId: 'caab0f18-e467-4d88-ae65-c08f769b8e01',
    displayName: 'Andrew',
    roles: ['owner', 'member']
  },
  {
    userId: 'a7530e0a-85f0-4f77-a617-2a5ab2a959ec',
    displayName: 'Dexter',
    roles: ['member']
  }
];

export const SampleSelfMember: AadUserConversationMember = {
  userId: 'a7530e0a-85f0-4f77-a617-2a5ab2a959ec',
  displayName: 'Dexter',
  roles: ['owner', 'member']
};

export const SampleChat: Chat = {
  id: '1698961635343:',
  topic: 'Dexter',
  chatType: 'oneOnOne',
  createdDateTime: '2021-07-06T21:58:36.141Z',
  lastUpdatedDateTime: '2021-07-06T21:58:36.141Z',
  lastMessagePreview: SampleMessagePreviewFromAndrew,
  members: SampleMembers
};

// update this to your own sample
export const SampleSelfChat: Chat = {
  id: '16989616353474',
  topic: 'SampleSelfChat',
  chatType: 'oneOnOne',
  createdDateTime: '2021-07-06T21:58:36.141Z',
  lastUpdatedDateTime: '2021-07-06T21:58:36.141Z',
  lastMessagePreview: SampleMessagePreviewFromDexter,
  members: [SampleSelfMember]
};

// use a static time (noon in this case) to avoid race conditions with the test
const today = new Date();
today.setHours(12, 0, 0, 0); // sets the time to 12:00:00.000
const todayISOString = today.toISOString();

export const SampleTodayChat: Chat = {
  id: '1698961635345:',
  topic: 'SampleTodayChat',
  chatType: 'oneOnOne',
  createdDateTime: '2021-07-06T21:58:36.141Z',
  lastUpdatedDateTime: todayISOString,
  lastMessagePreview: SampleMessagePreviewFromAndrew,
  members: SampleMembers
};

const yesterday = new Date();
yesterday.setDate(yesterday.getDate() - 1);
yesterday.setHours(12, 0, 0, 0);
const yesterdayISOString = yesterday.toISOString();

export const SampleYesterdayChat: Chat = {
  id: '1698961635346:',
  topic: 'SampleYesterdayChat',
  chatType: 'oneOnOne',
  createdDateTime: '2021-07-06T21:58:36.141Z',
  lastUpdatedDateTime: yesterdayISOString,
  lastMessagePreview: SampleMessagePreviewFromDexter,
  members: SampleMembers
};

/* section: GROUP SAMPLE DATA */

export const SampleGroupMembers: AadUserConversationMember[] = [
  {
    userId: 'a7530e0a-85f0-4f77-a617-2a5ab2a959ec',
    displayName: 'Dexter',
    roles: ['owner', 'member']
  },
  {
    userId: 'caab0f18-e467-4d88-ae65-c08f769b8e01',
    displayName: 'Andrew',
    roles: ['member']
  },
  {
    userId: '681c1ef5-c966-4124-9ea6-2ef2eb5ca3ba',
    displayName: 'Matthew',
    roles: ['member']
  },
  {
    userId: '4287e806-6904-41d5-a168-a42bf0411d46',
    displayName: 'Mark',
    roles: ['member']
  },
  {
    userId: 'ae6a7c20-e2c0-49e9-ac90-8629a073338c',
    displayName: 'Luke',
    roles: ['member']
  },
  {
    userId: '8f018b51-0f63-453b-88a5-8073e196b14b',
    displayName: 'Johnathan',
    roles: ['member']
  }
];

export const SampleGroupPreviewMessageFromSystem: ChatMessageInfo = {
  id: '1698961635347',
  createdDateTime: '2023-11-02T21:47:15.347Z',
  isDeleted: false,
  messageType: 'unknownFutureValue',
  body: {
    contentType: 'html',
    content: '<systemEventMessage/>'
  },
  eventDetail: {
    members: [
      {
        userId: '8d208569-5838-42d7-ac03-c798d8194a0a',
        displayName: 'Brandy',
        roles: ['member']
      }
    ] as AadUserConversationMember[],
    initiator: {
      user: {
        id: 'a7530e0a-85f0-4f77-a617-2a5ab2a959ec',
        displayName: 'Dexter'
      }
    }
  } as MembersAddedEventMessageDetail
};

export const SampleGroupPreviewMessageFromPerson: ChatMessageInfo = {
  id: '1698961635348',
  createdDateTime: '2023-12-13T23:27:24.405Z',
  isDeleted: false,
  messageType: 'message',
  body: {
    contentType: 'text',
    content: 'Hi everyone, call me Matt.'
  },
  from: {
    user: {
      id: '681c1ef5-c966-4124-9ea6-2ef2eb5ca3ba',
      displayName: 'Matthew'
    }
  } as ChatMessageFromIdentitySet
};

export const SampleGroupChat: Chat = {
  id: '1698961635349:',
  topic: 'Dexter, Andrew, Matthew, Mark, Luke, Johnathan',
  createdDateTime: '2021-07-06T21:58:36.141Z',
  lastUpdatedDateTime: '2021-07-06T21:58:36.141Z',
  chatType: 'group',
  lastMessagePreview: SampleGroupPreviewMessageFromPerson,
  members: SampleGroupMembers
};

export const SampleGroupChatMembershipChange: Chat = {
  id: '1698961635350:',
  topic: 'Group Chat With New Members',
  createdDateTime: '2023-11-02T21:46:58.373Z',
  lastUpdatedDateTime: '2023-11-02T21:47:15.292Z',
  chatType: 'group',
  lastMessagePreview: SampleGroupPreviewMessageFromSystem,
  members: null // note: incoming payload for membership changes does not include  existing members
};

export const SampleChats = {
  oneOnOne: { SampleChat, SampleSelfChat, SampleTodayChat, SampleYesterdayChat },
  group: { SampleGroupChat, SampleGroupChatMembershipChange }
};
