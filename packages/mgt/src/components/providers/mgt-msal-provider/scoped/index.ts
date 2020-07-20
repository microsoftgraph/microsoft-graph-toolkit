import { ScopedElementsMixin } from '@open-wc/scoped-elements';
import { ComponentRegistry } from '../../../../utils/ComponentRegistry';
import { MgtMsalProvider } from '../mgt-msal-provider';

ComponentRegistry.register(
  'mgt-msal-provider',
  class extends ScopedElementsMixin(MgtMsalProvider) {
    // tslint:disable-next-line: completed-docs
    static get scopedElements() {
      return ComponentRegistry.scopedElements;
    }
  }
);

export * from '../mgt-msal-provider';
