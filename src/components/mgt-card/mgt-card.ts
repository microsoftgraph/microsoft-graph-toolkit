import { css, customElement, property } from 'lit-element';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import { MgtTemplatedComponent } from '../templatedComponent';

import 'adaptivecards/dist/adaptivecards';

declare var AdaptiveCards: any;

@customElement('mgt-card')
export class MgtCard extends MgtTemplatedComponent {
  static get styles() {
    return css`
      .title {
        color: red;
      }
    `;
  }

  @property() public content: string;

  // assignment to this property will re-render the component
  @property() public hostConfig: any = new AdaptiveCards.HostConfig({
    fontFamily: 'Segoe UI, Helvetica Neue, sans-serif'
  });

  @property() public query: string;
  private _card: any; // need to add typing;

  private _lastQueryRendered: string;

  constructor() {
    super();
    this._card = new AdaptiveCards.AdaptiveCard();

    // Set the adaptive card's event handlers. onExecuteAction is invoked
    // whenever an action is clicked in the card
    this._card.onExecuteAction = action => this.onExecuteAction(action);
  }

  public attributeChangedCallback(name, oldval, newval) {
    super.attributeChangedCallback(name, oldval, newval);

    // TODO: handle when an attribute changes.
    //
    // Ex: load data when the name attribute changes
    if (name === 'query' && oldval !== newval) {
      this.loadData();
    }
  }

  public firstUpdated() {
    Providers.onProviderUpdated(() => this.loadData());
    this.loadData();
  }

  public render() {
    const root = this.renderRoot;

    while (root.firstChild) {
      root.removeChild(root.firstChild);
    }

    if (!this.content) {
      return null;
    }

    this._card.hostConfig = this.hostConfig;

    this._card.parse(this.content);
    const renderedCard = this._card.render();

    root.appendChild(renderedCard);
  }

  private async loadData() {
    const provider = Providers.globalProvider;

    if (!this.query || !provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    if (this._lastQueryRendered === this.query) {
      return;
    }

    this._lastQueryRendered = this.query;

    const client = provider.graph.client;
    const result = await client.api(this.query).get();

    if (result) {
      const response = await fetch('https://templates.adaptivecards.io/find', {
        body: JSON.stringify(result),
        method: 'post'
      });
      const template = await response.json();
      const templateUri = template[0].templateUrl;

      const contentResponse = await fetch('https://templates.adaptivecards.io/' + templateUri, {
        body: JSON.stringify(result),
        method: 'post'
      });

      this.content = await contentResponse.json();
    }
  }

  private onExecuteAction(action) {
    this.fireCustomEvent('action', action);
  }
}
