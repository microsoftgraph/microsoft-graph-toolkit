import { ScopedElementsMixin } from '@open-wc/scoped-elements';
import { ComponentRegistry } from '../../../utils/ComponentRegistry';
import { MgtAgenda } from '../mgt-agenda';

ComponentRegistry.register(
  'mgt-agenda',
  class extends ScopedElementsMixin(MgtAgenda) {
    // tslint:disable-next-line: completed-docs
    static get scopedElements() {
      return ComponentRegistry.scopedElements;
    }
  }
);

export * from '../mgt-agenda';
