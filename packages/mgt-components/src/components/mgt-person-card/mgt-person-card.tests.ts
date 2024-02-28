/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { expect } from '@open-wc/testing';
import { MockProvider, Providers } from '@microsoft/mgt-element';
import { MgtPersonCard, registerMgtPersonCardComponent } from './mgt-person-card';

describe('mgt-person - tests', () => {
  registerMgtPersonCardComponent();
  Providers.globalProvider = new MockProvider(true);

  it('should provide required scopes', () => {
    const expectedScopes = [
      'User.Read.All',
      'People.Read.All',
      'Sites.Read.All',
      'Mail.Read',
      'Mail.ReadBasic',
      'Contacts.Read',
      'Chat.ReadWrite'
    ];
    const scopes = MgtPersonCard.requiredScopes;
    // eslint-disable-next-line @typescript-eslint/no-unused-expressions
    expect(scopes).to.be.not.undefined;
    expect(scopes).to.have.members(expectedScopes);
  });
});
