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

        this.people = (await client.contacts()).slice(0, this.showMax);
      }
    }
  }

  render() {
    let templates = this.getTemplates();
    this.removeSlottedElements();

    if (this.people) {
      if (templates['default']) {
        return this.renderTemplate(templates['default'], { people: this.people }, 'global');
      } else {
        return html`
          <ul class="people-list">
            ${this.people.map(
              person =>
                html`
                  <li>
                    ${templates['person']
                      ? this.renderTemplate(templates['person'], { person: person }, person.id)
                      : this.renderPerson(person)}
                  </li>
                `
            )}
          </ul>
        `;
      }
    } else {
      return templates['no-data'] ? this.renderTemplate(templates['no-data'], null, 'no-data') : html``;
    }
  }

  private renderPerson(person: MicrosoftGraph.Person) {
    return html`
      <div class="people-person">
        <mgt-person person-details=${JSON.stringify(person)}></mgt-person>
      </div>
    `;
  }
}
