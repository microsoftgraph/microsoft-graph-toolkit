/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, LitElement, property } from 'lit-element';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import '../../styles/fabric-icon-font';
import '../mgt-person/mgt-person';
import { MgtTemplatedComponent } from '../templatedComponent';
import { styles } from './mgt-people-css';

@customElement('mgt-people')
export class MgtPeople extends MgtTemplatedComponent {
  /* TODO: Do we want a query property for loading groups from calls? */

  static get styles() {
    return styles;
  }

  @property({
    attribute: 'people',
    type: Object
  })
  public people: Array<MicrosoftGraph.User | MicrosoftGraph.Person | MicrosoftGraph.Contact> = null;

  @property({
    attribute: 'show-max',
    type: Number
  })
  public showMax: number = 3;
  private _firstUpdated = false;

  constructor() {
    super();
  }

  public firstUpdated() {
    this._firstUpdated = true;
    Providers.onProviderUpdated(() => this.loadPeople());
    this.loadPeople();
  }

  public render() {
    if (this.people) {
      return (
        this.renderTemplate('default', { people: this.people }) ||
        html`
          <ul class="people-list">
            ${this.people.slice(0, this.showMax).map(
              person =>
                html`
                  <li class="people-person">
                    ${this.renderTemplate('person', { person }, person.displayName) || this.renderPerson(person)}
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

  private async loadPeople() {
    if (!this._firstUpdated) {
      return;
    }

    if (!this.people) {
      const provider = Providers.globalProvider;

      if (provider && provider.state === ProviderState.SignedIn) {
        const client = Providers.globalProvider.graph;

        this.people = (await client.getPeople()).slice(0, this.showMax);
      }
    }
  }

  private renderPerson(person: MicrosoftGraph.Person) {
    // set image to @ to flag the mgt-person component to
    // query the image from the graph
    return html`
      <mgt-person .personDetails=${person} .personImage=${'@'}></mgt-person>
    `;
  }
}
