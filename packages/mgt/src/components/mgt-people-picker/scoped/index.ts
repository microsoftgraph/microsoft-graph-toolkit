import { ScopedElementsMixin } from '@open-wc/scoped-elements';
import { ComponentRegistry } from '../../../utils/ComponentRegistry';
import { MgtPeoplePicker } from '../mgt-people-picker';

ComponentRegistry.register(
  'mgt-people-picker',
  class extends ScopedElementsMixin(MgtPeoplePicker) {
    // tslint:disable-next-line: completed-docs
    static get scopedElements() {
      return ComponentRegistry.scopedElements;
    }
  }
);

export * from '../mgt-people-picker';
