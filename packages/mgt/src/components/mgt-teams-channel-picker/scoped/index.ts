import { ScopedElementsMixin } from '@open-wc/scoped-elements';
import { ComponentRegistry } from '../../../utils/ComponentRegistry';
import { MgtTeamsChannelPicker } from '../mgt-teams-channel-picker';

ComponentRegistry.register(
  'mgt-teams-channel-picker',
  class extends ScopedElementsMixin(MgtTeamsChannelPicker) {
    // tslint:disable-next-line: completed-docs
    static get scopedElements() {
      return ComponentRegistry.scopedElements;
    }
  }
);

export * from '../mgt-teams-channel-picker';
