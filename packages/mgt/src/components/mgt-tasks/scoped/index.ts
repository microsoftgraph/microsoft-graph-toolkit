import { ScopedElementsMixin } from '@open-wc/scoped-elements';
import { ComponentRegistry } from '../../../utils/ComponentRegistry';
import { MgtTasks } from '../mgt-tasks';

ComponentRegistry.register(
  'mgt-tasks',
  class extends ScopedElementsMixin(MgtTasks) {
    // tslint:disable-next-line: completed-docs
    static get scopedElements() {
      return ComponentRegistry.scopedElements;
    }
  }
);

export * from '../mgt-tasks';
