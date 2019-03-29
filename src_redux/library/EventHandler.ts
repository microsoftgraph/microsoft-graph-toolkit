export type EventHandler<E> = (event: E) => void;
export class EventDispatcher<E> {
  private eventHandlers: EventHandler<E>[] = [];
  fire(event: E) {
    for (let handler of this.eventHandlers) {
      handler(event);
    }
  }
  register(eventHandler: EventHandler<E>) {
    this.eventHandlers.push(eventHandler);
  }
}
