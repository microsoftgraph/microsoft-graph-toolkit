/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { chatOperationScopes } from '../statefulClient/chatOperationScopes';
import { reduceScopes } from './reduceScopes';

describe('reduceScopes tests', () => {
  it('should choose the first entry in all arrays when they match', () => {
    const scopeSet = [
      ['chat.read', 'chat.write'],
      ['chat.read', 'chat.read.all'],
      ['chat.read', 'chat.write']
    ];
    const actual = reduceScopes(scopeSet);
    expect(actual).toEqual(['chat.read']);
  });

  it('should choose chat.read and chat.read.all', () => {
    const scopeSet = [['chat.read', 'chat.write'], ['chat.read.all'], ['chat.read', 'chat.write']];
    const actual = reduceScopes(scopeSet);
    expect(actual).toEqual(['chat.read', 'chat.read.all']);
  });

  it('should choose chat.write and chat.read', () => {
    const scopeSet = [
      ['chat.read', 'chat.write'],
      ['chat.write', 'chat.read.all'],
      ['chat.read', 'chat.write']
    ];
    const actual = reduceScopes(scopeSet);
    expect(actual).toEqual(['chat.write']);
  });

  it('should pass the chat gpt test', () => {
    const input = [
      ['apple', 'banana', 'cherry'],
      ['dog', 'cat', 'frog'],
      ['red', 'cherry', 'blue']
    ];
    const actual = reduceScopes(input);
    expect(actual).toEqual(['cherry', 'dog']);
  });

  it('will slightly over calculate for complex sets', () => {
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
    expect(actual).toEqual(['user.read', 'user.readbasic', 'user.write']);
  });

  it('will provide the correct minimal permissions for chat', () => {
    const input = Object.keys(chatOperationScopes).map(k => chatOperationScopes[k]);
    const actual = reduceScopes(input);
    // NOTE: the idea result here is ['user.read', 'user.write']
    // however the given result results in a higher scope count and not higher effective permissions
    expect(actual).toEqual(['Chat.ReadWrite', 'ChatMember.ReadWrite']);
  });
});
