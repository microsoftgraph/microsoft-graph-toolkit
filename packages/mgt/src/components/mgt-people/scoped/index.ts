import { ScopedElementsMixin } from '@open-wc/scoped-elements';
import { ComponentRegistry } from '../../../utils/ComponentRegistry';
import { MgtPeople } from '../mgt-people';

ComponentRegistry.register(
  'mgt-people',
  class extends ScopedElementsMixin(MgtPeople) {
    // tslint:disable-next-line: completed-docs
    static get scopedElements() {
      return ComponentRegistry.scopedElements;
    }
  }
);

export * from '../mgt-people';
