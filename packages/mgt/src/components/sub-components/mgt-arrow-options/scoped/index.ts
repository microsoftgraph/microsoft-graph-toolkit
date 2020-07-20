import { ScopedElementsMixin } from '@open-wc/scoped-elements';
import { ComponentRegistry } from '../../../../utils/ComponentRegistry';
import { MgtArrowOptions } from '../mgt-arrow-options';

ComponentRegistry.register(
  'mgt-arrow-options',
  class extends ScopedElementsMixin(MgtArrowOptions) {
    // tslint:disable-next-line: completed-docs
    static get scopedElements() {
      return ComponentRegistry.scopedElements;
    }
  }
);

export * from '../mgt-arrow-options';
