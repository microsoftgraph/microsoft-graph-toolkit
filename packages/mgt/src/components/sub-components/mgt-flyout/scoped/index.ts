import { ScopedElementsMixin } from '@open-wc/scoped-elements';
import { ComponentRegistry } from '../../../../utils/ComponentRegistry';
import { MgtFlyout } from '../mgt-flyout';

ComponentRegistry.register(
  'mgt-flyout',
  class extends ScopedElementsMixin(MgtFlyout) {
    // tslint:disable-next-line: completed-docs
    static get scopedElements() {
      return ComponentRegistry.scopedElements;
    }
  }
);

export * from '../mgt-flyout';
