/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * An EventHandler for custom events
 */
export type EventHandler<E> = (event: E) => void;

/**
 * Provider EventDispatcher
 *
 * @export
 * @class EventDispatcher
 * @template E
 */
export class EventDispatcher<E> {
  private eventHandlers: Array<EventHandler<E>> = [];
  /**
   * fires event handler
   *
   * @param {E} event
   * @memberof EventDispatcher
   */

  public fire(event: E) {
    for (const handler of this.eventHandlers) {
      handler(event);
    }
  }

  /**
   * adds eventHandler
   *
   * @param {EventHandler<E>} eventHandler
   * @memberof EventDispatcher
   */
  public add(eventHandler: EventHandler<E>) {
    this.eventHandlers.push(eventHandler);
  }

  /**
   * removes eventHandler
   *
   * @param {EventHandler<E>} eventHandler
   * @memberof EventDispatcher
   */
  public remove(eventHandler: EventHandler<E>) {
    for (let i = 0; i < this.eventHandlers.length; i++) {
      if (this.eventHandlers[i] === eventHandler) {
        this.eventHandlers.splice(i, 1);
        i--;
      }
    }
  }
}
