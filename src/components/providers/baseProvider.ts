import { Providers } from '../../Providers';
import { MgtBaseComponent } from '../baseComponent';

export abstract class MgtBaseProvider extends MgtBaseComponent {
  constructor() {
    super();
    Providers.onProviderUpdated(() => this.loadState());
    this.loadState();
  }

  private async loadState() {
    const provider = Providers.globalProvider;

    if (provider) {
      provider.onStateChanged(() => {
        this.fireCustomEvent('onStateChanged', provider.state);
      });
    }
  }
}
