/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { LitElement, html, customElement, property } from 'lit-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import { styles } from './mgt-people-css';

import '../mgt-person/mgt-person';
import '../../styles/fabric-icon-font';
import { MgtTemplatedComponent } from '../templatedComponent';
import { MgtPersonDetails } from '../mgt-person/mgt-person';

@customElement('mgt-people')
export class MgtPeople extends MgtTemplatedComponent {
  private _firstUpdated = false;
  /**
   * people property, containing array of people used in the component.
   * @type {Array<MgtPersonDetails>}
   */
  @property({
    attribute: 'people',
    type: Object
  })
  people: Array<MgtPersonDetails> = null;

  /**
   * show-max property, developer determined max people shown in component
   * @type {number}
   */
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
  }

  firstUpdated() {
    this._firstUpdated = true;
    Providers.onProviderUpdated(() => this.loadPeople());
    this.loadPeople();
  }

  private async loadPeople() {
    if (!this._firstUpdated) {
      return;
    }

    if (!this.people) {
      let provider = Providers.globalProvider;

      if (provider && provider.state === ProviderState.SignedIn) {
        let client = Providers.globalProvider.graph;

        this.people = (await client.getPeople()).slice(0, this.showMax);
        for (let person of this.people) {
          // set image to @ to flag the mgt-person component to
          // query the image from the graph
          person.image = '@';
        }
      }
    }
  }

  render() {
    if (this.people) {
      return (
        this.renderTemplate('default', { people: this.people }) ||
        html`
          <ul class="people-list">
            ${this.people.slice(0, this.showMax).map(
              person =>
                html`
                  <li class="people-person">
                    ${this.renderTemplate('person', { person: person }, person.displayName) ||
                      this.renderPerson(person)}
                  </li>
                `
            )}
            ${this.people.length > this.showMax
              ? this.renderTemplate('overflow', {
                  people: this.people,
                  max: this.showMax,
                  extra: this.people.length - this.showMax
                }) ||
                html`
                  <li>+${this.people.length - this.showMax}</li>
                `
              : null}
          </ul>
        `
      );
    } else {
      return this.renderTemplate('no-data', null) || html``;
    }
  }

  private renderPerson(person: MicrosoftGraph.Person) {
    return html`
      <mgt-person .personDetails=${person}></mgt-person>
    `;
  }
}
