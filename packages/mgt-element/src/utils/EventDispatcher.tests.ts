/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { EventDispatcher } from './EventDispatcher';
import { assert, restore, fake } from 'sinon';

describe('EventDispatcher tests', () => {
  afterEach(() => {
    // Restore the default sandbox here
    restore();
  });
  it('should add and remove event handlers', () => {
    const dispatcher = new EventDispatcher();
    const handler1 = fake();
    const handler2 = fake();
    dispatcher.add(handler1);
    dispatcher.add(handler2);
    dispatcher.fire('event');

    assert.calledOnce(handler1);
    assert.calledOnce(handler2);
    dispatcher.remove(handler1);
    dispatcher.fire('event');
    assert.calledOnce(handler1);
    assert.callCount(handler2, 2);
  });

  it('should not throw when remove is called with an unregistered handler', () => {
    try {
      const dispatcher = new EventDispatcher();
      const handler1 = fake();
      dispatcher.remove(handler1);
    } catch (e) {
      assert.fail('should not throw');
    }
    assert.pass('did not throw');
  });
});
