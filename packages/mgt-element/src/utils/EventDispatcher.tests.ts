import { it } from '@jest/globals';
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
});
