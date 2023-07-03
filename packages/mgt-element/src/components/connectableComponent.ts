import { state } from 'lit-element';
import { MgtTemplatedComponent } from './templatedComponent';
import { EventHandler } from '../utils/EventDispatcher';
import { isObjectLike } from 'lodash-es';
import { IComponentBinding } from '../utils/IComponentBinding';
import { LocalizationHelper } from '../utils/LocalizationHelper';
import { ILocalizedString } from '../utils/ILocalizedString';

export abstract class MgtConnectableComponent extends MgtTemplatedComponent {
  /**
   * Flag indicating if data have been rendered at least once
   */
  @state()
  renderedOnce = false;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  protected declare _eventHanlders: Map<
    string,
    { component: HTMLElement; event: string; handler: EventHandler<Event> }
  >;

  constructor() {
    super();

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    this._eventHanlders = new Map<string, { component: HTMLElement; event: string; handler: EventHandler<Event> }>();
  }

  public disconnectedCallback(): void {
    super.disconnectedCallback();

    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    for (const [key, value] of this._eventHanlders.entries()) {
      value.component.removeEventListener(value.event, value.handler);
    }
  }

  protected bindComponents(bindings: IComponentBinding[]) {
    const bindElement = (componentElement, binding) => {
      this._eventHanlders.set(`${binding.id}-${binding.eventName}`, {
        component: componentElement,
        event: binding.eventName,
        handler: binding.callbackFunction
      });
      componentElement.addEventListener(binding.eventName, binding.callbackFunction);
    };

    bindings.forEach(binding => {
      const eventKey = `${binding.id}-${binding.eventName}`;

      // The components should be already available in the DOM as they are statically defined in the page
      const componentElement = document.getElementById(binding.id);

      if (componentElement && !this._eventHanlders.get(eventKey)) {
        bindElement(componentElement, binding);
      }
    });
  }

  protected getLocalizedString(string: ILocalizedString | string) {
    if (isObjectLike(string)) {
      const localizedValue = string[LocalizationHelper.strings?.language];

      if (!localizedValue) {
        const defaultLabel = (string as ILocalizedString).default;
        return defaultLabel ? defaultLabel : '{{Translation not found}}';
      }

      return localizedValue;
    }

    return string;
  }
}
