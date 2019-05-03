import { LitElement, html, customElement, property } from 'lit-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import { styles } from './mgt-people-css';

import '../mgt-person/mgt-person';
import '../../styles/fabric-icon-font';
import { MgtTemplatedComponent } from '../templatedComponent';

@customElement('mgt-people')
export class MgtPeople extends MgtTemplatedComponent {
  private _peopleData: Array<MicrosoftGraph.Person> = null;

  @property({
    attribute: 'people',
    type: Array
  })
  people: Array<MicrosoftGraph.Person> = null;

  @property({
    attribute: 'show-max',
    type: Number
  })
  showMax: number = 3;

  /* TODO: Do we want a query property for loading groups from calls? */

  static get styles() {
    return styles;
  }

  constructor() {
    super();
    Providers.onProviderUpdated(() => this.loadPeople());
    this.loadPeople();
  }

  private async loadPeople() {
    if (!this.people) {
      let provider = Providers.globalProvider;

      if (provider && provider.state === ProviderState.SignedIn) {
        let client = Providers.globalProvider.graph;

        this._peopleData = await client.getPeople();
        this.people = this._peopleData.slice(0, this.showMax);
      }
    }
  }

  attributeChangedCallback(name, oldValue, newValue) {
    super.attributeChangedCallback(name, oldValue, newValue);

    if (this._peopleData && name === 'show-max') {
      this.people = this._peopleData.slice(0, this.showMax);
    }
  }

  render() {
    if (this.people) {
      return (
        this.renderTemplate('default', { people: this.people }) ||
        html`
          <ul class="people-list">
            ${this.people.map(
              person =>
                html`
                  <li>
                    ${this.renderTemplate('person', { person: person }, person.id) || this.renderPerson(person)}
                  </li>
                `
            )}
          </ul>
        `
      );
    } else {
      return this.renderTemplate('no-data', null) || html``;
    }
  }

  private renderPerson(person: MicrosoftGraph.Person) {
    return html`
      <div class="people-person">
        <mgt-person person-query=${person.userPrincipalName}></mgt-person>
      </div>
    `;
  }
}
