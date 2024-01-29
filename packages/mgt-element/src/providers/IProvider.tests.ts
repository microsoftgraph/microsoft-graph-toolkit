/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IProvider } from './IProvider';
import { SimpleProvider } from './SimpleProvider';
import { expect } from '@open-wc/testing';

describe('IProvider.needsAdditionalScopes tests', () => {
  let p: IProvider;
  beforeEach(() => {
    p = new SimpleProvider(
      () => Promise.resolve('fake-token'),
      () => Promise.resolve(),
      () => Promise.resolve()
    );
    p.approvedScopes = ['user.read', 'group.read.all', 'presence.read'];
  });
  it('should provide an empty array when one scope is already present', async () => {
    const result = p.needsAdditionalScopes(['groupmember.read.all', 'group.read.all']);
    await expect(result).to.eql([]);
  });
  it('should provide an empty array when one scope is already present ignoring case of scopes in provider', async () => {
    p.approvedScopes = ['user.read', 'Group.Read.All', 'presence.read'];
    const result = p.needsAdditionalScopes(['groupmember.read.all', 'group.read.all']);
    await expect(result).to.eql([]);
  });
  it('should provide an empty array when one scope is already present ignoring case of scopes in provider', async () => {
    const result = p.needsAdditionalScopes(['groupmember.read.all', 'Group.Read.All']);
    await expect(result).to.eql([]);
  });
  it('should provide an the first element in the passed array where there is no overlap', async () => {
    const result = p.needsAdditionalScopes(['groupmember.read.all', 'group.readwrite.all']);
    await expect(result).to.eql(['groupmember.read.all']);
  });
});
