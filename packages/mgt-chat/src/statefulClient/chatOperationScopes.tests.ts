/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

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
