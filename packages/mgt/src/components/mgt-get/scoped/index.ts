import { ScopedElementsMixin } from '@open-wc/scoped-elements';
import { ComponentRegistry } from '../../../utils/ComponentRegistry';
import { MgtGet } from '../mgt-get';

ComponentRegistry.register(
  'mgt-get',
  class extends ScopedElementsMixin(MgtGet) {
    // tslint:disable-next-line: completed-docs
    static get scopedElements() {
      return ComponentRegistry.scopedElements;
    }
  }
);

export * from '../mgt-get';
