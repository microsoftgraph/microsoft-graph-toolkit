import { ScopedElementsMixin } from '@open-wc/scoped-elements';
import { ComponentRegistry } from '../../../../utils/ComponentRegistry';
import { MgtProxyProvider } from '../mgt-proxy-provider';

ComponentRegistry.register(
  'mgt-proxy-provider',
  class extends ScopedElementsMixin(MgtProxyProvider) {
    // tslint:disable-next-line: completed-docs
    static get scopedElements() {
      return ComponentRegistry.scopedElements;
    }
  }
);

export * from '../mgt-proxy-provider';
