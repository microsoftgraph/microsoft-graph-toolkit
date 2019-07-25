import { html, css, customElement, property } from 'lit-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { MgtTemplatedComponent } from '../templatedComponent';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import { styles } from './mgt-person-card-css';

@customElement('mgt-person-card')
export class MgtPersonCard extends MgtTemplatedComponent {
  @property({
    attribute: 'my-title',
    type: String
  })
  myTitle: string = 'My First Component';

  // assignment to this property will re-render the component
  @property({ attribute: false }) private _me: MicrosoftGraph.User;

  attributeChangedCallback(name, oldval, newval) {
    super.attributeChangedCallback(name, oldval, newval);

    // TODO: handle when an attribute changes.
    //
    // Ex: load data when the name attribute changes
    // if (name === 'person-id' && oldval !== newval){
    //  this.loadData();
    // }
  }

  firstUpdated() {
    Providers.onProviderUpdated(() => this.loadData());
    this.loadData();
  }

  private async loadData() {
    let provider = Providers.globalProvider;

    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    // TODO: load data from the graph
    this._me = await provider.graph.client
      .api('/me')
      .version('beta')
      .get();
    console.log(this._me);
  }

  static get styles() {
    return styles;
  }

  render() {
    if (this._me) {
      const user = this._me;

      let phone, department, jobTitle, email, location;

      if (user.businessPhones && user.businessPhones.length > 0) {
        phone = html`
          <div class="phone subtitle">${user.businessPhones[0]}</div>
        `;
      }

      if (user.department) {
        department = html`
          <div><b>department: </b>${user.department}</div>
        `;
      }

      if (user.jobTitle) {
        jobTitle = html`
          <div><b>job title: </b>${user.jobTitle}</div>
        `;
      }

      if (user.mail) {
        email = html`
          <div class="email subtitle">${user.mail}</div>
        `;
      }

      if (user.officeLocation) {
        location = html`
          <div><b>office location: </b>${user.officeLocation}</div>
        `;
      }

      return html`
        <div class="root">
          <div class="displayname">${user.displayName}</div>
          ${phone} ${department} ${jobTitle} ${email} ${location}
        </div>
      `;
    }
  }
}
