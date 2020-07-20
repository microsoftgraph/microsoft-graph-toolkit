import { ScopedElementsMixin } from '@open-wc/scoped-elements';
import { ComponentRegistry } from '../../../utils/ComponentRegistry';
import { MgtPersonCard } from '../mgt-person-card';

ComponentRegistry.register(
  'mgt-person-card',
  class extends ScopedElementsMixin(MgtPersonCard) {
    // tslint:disable-next-line: completed-docs
    static get scopedElements() {
      return ComponentRegistry.scopedElements;
    }
  }
);

export * from '../mgt-person-card';
