import { LitElement, html, customElement, property } from 'lit-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import { styles } from './mgt-people-css';

import '../mgt-person/mgt-person';
import '../../styles/fabric-icon-font';

@customElement('mgt-people')
export class MgtPeople extends LitElement {
  @property({
    attribute: 'people',
    type: Array
  })
  people: [];

  /* TODO: Do we want a query property for loading groups from calls? */

  // @property() listTemplateFunction: (group: any) => string;
  @property() personTemplateFunction: (people: any) => string;

  static get styles() {
    return styles;
  }

  constructor() {
    super();
  }

  render() {
    if (this.people) {
      // remove slotted elements inserted initially
      while (this.lastChild) {
        this.removeChild(this.lastChild);
      }

      return html`
        <ul class="people-list">
          ${this.people.map(
            person =>
              html`
                <li>
                  ${this.personTemplateFunction ? this.renderPersonTemplate(person) : this.renderPerson(person)}
                </li>
              `
          )}
        </ul>
      `;
    } else {
      return html`
        <div></div>
      `;
    }
  }

  private renderPerson(person: MicrosoftGraph.Person) {
    return html`
      <div class="people-person">
        <mgt-person person-details=${JSON.stringify(person)} show-name></mgt-person>
      </div>
    `;
  }

  private renderPersonTemplate(person: MicrosoftGraph.Person) {
    let content: any = this.personTemplateFunction(person);
    if (typeof content === 'string') {
      return html`
        <div>${content}</div>
      `;
    } else {
      let div = document.createElement('div');
      div.slot = person.displayName;
      div.appendChild(content);

      this.appendChild(div);
      return html`
        <slot name=${person.displayName}></slot>
      `;
    }
  }
}
