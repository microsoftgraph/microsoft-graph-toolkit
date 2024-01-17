/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { chatOperationScopes } from '../statefulClient/chatOperationScopes';
import { reduceScopes } from './reduceScopes';
import { expect } from '@open-wc/testing';

describe('reduceScopes tests', () => {
  it('should choose the first entry in all arrays when they match', async () => {
    const scopeSet = [
      ['chat.read', 'chat.write'],
      ['chat.read', 'chat.read.all'],
      ['chat.read', 'chat.write']
    ];
    const actual = reduceScopes(scopeSet);
    await expect(actual).to.eql(['chat.read']);
  });

  it('should choose chat.read and chat.read.all', async () => {
    const scopeSet = [['chat.read', 'chat.write'], ['chat.read.all'], ['chat.read', 'chat.write']];
    const actual = reduceScopes(scopeSet);
    await expect(actual).to.eql(['chat.read', 'chat.read.all']);
  });

  it('should choose chat.write and chat.read', async () => {
    const scopeSet = [
      ['chat.read', 'chat.write'],
      ['chat.write', 'chat.read.all'],
      ['chat.read', 'chat.write']
    ];
    const actual = reduceScopes(scopeSet);
    await expect(actual).to.eql(['chat.write']);
  });

  it('should pass the chat gpt test', async () => {
    const input = [
      ['apple', 'banana', 'cherry'],
      ['dog', 'cat', 'frog'],
      ['red', 'cherry', 'blue']
    ];
    const actual = reduceScopes(input);
    await expect(actual).to.eql(['cherry', 'dog']);
  });

  it('will slightly over calculate for complex sets', async () => {
    const input = [
      ['user.read', 'user.write'],
      ['user.readbasic', 'user.read'],
      ['basic', 'user.readbasic', 'user.read.all', 'user.write'],
      ['user.readbasic', 'user.readbasic.all', 'user.read'],
      ['user.readbasic', 'user.read'],
      ['user.write']
    ];
    const actual = reduceScopes(input);
    // NOTE: the idea result here is ['user.read', 'user.write']
    // however the given result results in a higher scope count and not higher effective permissions
    await expect(actual).to.eql(['user.read', 'user.readbasic', 'user.write']);
  });

  it('will provide the correct minimal permissions for chat', async () => {
    const input = Object.keys(chatOperationScopes).map(k => chatOperationScopes[k]);
    const actual = reduceScopes(input);
    // NOTE: the idea result here is ['user.read', 'user.write']
    // however the given result results in a higher scope count and not higher effective permissions
    await expect(actual).to.eql(['Chat.ReadWrite', 'ChatMember.ReadWrite', 'TeamsAppInstallation.ReadForChat']);
  });
});
