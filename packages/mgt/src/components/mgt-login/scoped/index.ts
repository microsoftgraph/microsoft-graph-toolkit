import { ScopedElementsMixin } from '@open-wc/scoped-elements';
import { ComponentRegistry } from '../../../utils/ComponentRegistry';
import { MgtLogin } from '../mgt-login';

ComponentRegistry.register(
  'mgt-login',
  class extends ScopedElementsMixin(MgtLogin) {
    // tslint:disable-next-line: completed-docs
    static get scopedElements() {
      return ComponentRegistry.scopedElements;
    }
  }
);

export * from '../mgt-login';
