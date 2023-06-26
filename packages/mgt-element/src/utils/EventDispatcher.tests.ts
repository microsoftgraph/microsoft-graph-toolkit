/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { it } from '@jest/globals';
import { assert } from 'console';
import { EventDispatcher } from './EventDispatcher';

describe('EventDispatcher tests', () => {
  it('should add and remove event handlers', () => {
    const dispatcher = new EventDispatcher();
    const handler1 = jest.fn();
    const handler2 = jest.fn();
    dispatcher.add(handler1);
    dispatcher.add(handler2);
    dispatcher.fire('event');
    expect(handler1).toHaveBeenCalledTimes(1);
    expect(handler2).toHaveBeenCalledTimes(1);
    dispatcher.remove(handler1);
    dispatcher.fire('event');
    expect(handler1).toHaveBeenCalledTimes(1);
    expect(handler2).toHaveBeenCalledTimes(2);
  });
  it('should not throw when remove is called with an unregistered handler', () => {
    try {
      const dispatcher = new EventDispatcher();
      const handler1 = jest.fn();
      dispatcher.remove(handler1);
    } catch (e) {
      assert(false, 'should not throw');
    }
    assert(true, 'did not throw');
  });
});
