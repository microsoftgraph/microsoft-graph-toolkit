import { LitElement, customElement, property } from 'lit-element';
import { Providers } from '../../Providers';
import { WamProvider } from '../../providers/WamProvider';

@customElement('mgt-wam-provider')
export class MgtWamProvider extends LitElement {
  @property({ attribute: 'client-id' }) clientId: string;

  @property({ attribute: 'authority' }) authority?: string;

  firstUpdated(changedProperties) {
    this.validateAuthProps();
  }

  private validateAuthProps() {
    if (this.clientId !== undefined) {
      Providers.globalProvider = new WamProvider(this.clientId, this.authority);
    }
  }
}
