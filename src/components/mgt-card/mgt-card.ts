import { html, css, customElement, property } from 'lit-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { MgtTemplatedComponent } from '../templatedComponent';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import 'adaptivecards/dist/adaptivecards';

declare var AdaptiveCards: any;

@customElement('mgt-card')
export class MgtCard extends MgtTemplatedComponent {
  private _card: any; // need to add typing;

  @property() content: string;

  // assignment to this property will re-render the component
  @property() hostConfig: any = new AdaptiveCards.HostConfig({
    fontFamily: 'Segoe UI, Helvetica Neue, sans-serif'
  });

  @property() query: string;

  private _lastQueryRendered: string;

  constructor() {
    super();
    this._card = new AdaptiveCards.AdaptiveCard();

    // Set the adaptive card's event handlers. onExecuteAction is invoked
    // whenever an action is clicked in the card
    this._card.onExecuteAction = action => this.onExecuteAction(action);
  }

  attributeChangedCallback(name, oldval, newval) {
    super.attributeChangedCallback(name, oldval, newval);

    // TODO: handle when an attribute changes.
    //
    // Ex: load data when the name attribute changes
    if (name === 'query' && oldval !== newval) {
      this.loadData();
    }
  }

  private async loadData() {
    const provider = Providers.globalProvider;

    if (this.query == null || this.query == '' || !provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    if (this._lastQueryRendered === this.query) {
      return;
    }

    this._lastQueryRendered = this.query;

    const client = provider.graph.client;
    let result = await client.api(this.query).get();

    if (result) {
      let response = await fetch('https://templates.adaptivecards.io/find', {
        method: 'post',
        body: JSON.stringify(result)
      });
      let template = await response.json();
      let templateUri = template[0].templateUrl;

      let contentResponse = await fetch('https://templates.adaptivecards.io/' + templateUri, {
        method: 'post',
        body: JSON.stringify(result)
      });

      this.content = await contentResponse.json();
    }
  }

  private onExecuteAction(action) {
    this.fireCustomEvent('action', action);
  }

  firstUpdated() {
    Providers.onProviderUpdated(() => this.loadData());
    this.loadData();
  }

  static get styles() {
    return css`
      .title {
        color: red;
      }
    `;
  }

  render() {
    this._card.hostConfig = this.hostConfig;

    this._card.parse(this.content);
    let renderedCard = this._card.render();

    let root = this.renderRoot;

    while (root.firstChild) {
      root.removeChild(root.firstChild);
    }

    root.appendChild(renderedCard);
  }
}
