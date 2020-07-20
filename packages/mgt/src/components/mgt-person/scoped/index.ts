import { ScopedElementsMixin } from '@open-wc/scoped-elements';
import { ComponentRegistry } from '../../../utils/ComponentRegistry';
import { MgtPerson } from '../mgt-person';

ComponentRegistry.register(
  'mgt-person',
  class extends ScopedElementsMixin(MgtPerson) {
    // tslint:disable-next-line: completed-docs
    static get scopedElements() {
      return ComponentRegistry.scopedElements;
    }
  }
);

export * from '../mgt-person';
