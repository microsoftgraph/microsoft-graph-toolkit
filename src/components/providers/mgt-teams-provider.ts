import { LitElement, customElement, property } from 'lit-element';
import { Providers } from '../../Providers';
import { TeamsProvider } from '../../providers/TeamsProvider';

@customElement('mgt-teams-provider')
export class MgtTeamsProvider extends LitElement {
  private _provider: TeamsProvider;

  @property({
    type: String,
    attribute: 'client-id'
  })
  clientId = '';

  @property({
    type: String,
    attribute: 'login-popup-url'
  })
  loginPopupUrl = '';

  async firstUpdated(changedProperties) {
    this.validateAuthProps();
    if (await TeamsProvider.isAvailable()) {
      Providers.globalProvider = this._provider;
    }
  }

  private validateAuthProps() {
    if (this.clientId && this.loginPopupUrl) {
      if (!this._provider) {
        this._provider = new TeamsProvider(this.clientId, this.loginPopupUrl);
      }
      this._provider.clientId = this.clientId;
      this._provider.loginPopupUrl = this.loginPopupUrl;
    }
  }
}
