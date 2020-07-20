import { ScopedElementsMixin } from '@open-wc/scoped-elements';
import { ComponentRegistry } from '../../../../utils/ComponentRegistry';
import { MgtTeamsProvider } from '../mgt-teams-provider';

ComponentRegistry.register(
  'mgt-teams-provider',
  class extends ScopedElementsMixin(MgtTeamsProvider) {
    // tslint:disable-next-line: completed-docs
    static get scopedElements() {
      return ComponentRegistry.scopedElements;
    }
  }
);

export * from '../mgt-teams-provider';
