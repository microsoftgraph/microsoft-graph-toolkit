import { expect } from '@open-wc/testing';
import { allChatScopes } from './chatOperationScopes';

describe('chatOperationScopes tests', () => {
  it('should have a minimal permission set', () => {
    const expectedScopes = [
      'User.Read',
      'User.ReadBasic.All',
      'User.Read.All',
      'People.Read',
      'People.Read.All',
      'Presence.Read.All',
      'Sites.Read.All',
      'Mail.Read',
      'Mail.ReadBasic',
      'Contacts.Read',
      'Chat.ReadWrite',
      'ChatMember.ReadWrite'
    ];
    expect(allChatScopes).to.have.members(expectedScopes);
  });
});
