export interface IComponentBinding {
  /**
   * The DOM element ID to bind
   */
  id: string;

  /**
   * The event name to bind
   */
  eventName: string;

  /**
   * Function to call when the event is fired
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  callbackFunction: (e: CustomEvent<any>) => void;
}
