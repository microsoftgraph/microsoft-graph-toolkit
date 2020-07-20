import { ScopedElementsMixin } from '@open-wc/scoped-elements';
import { ComponentRegistry } from '../../../../utils/ComponentRegistry';
import { MgtDotOptions } from '../mgt-dot-options';

ComponentRegistry.register(
  'mgt-dot-options',
  class extends ScopedElementsMixin(MgtDotOptions) {
    // tslint:disable-next-line: completed-docs
    static get scopedElements() {
      return ComponentRegistry.scopedElements;
    }
  }
);

export * from '../mgt-dot-options';
