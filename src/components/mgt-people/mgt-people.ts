/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property } from 'lit-element';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import '../../styles/fabric-icon-font';
import '../mgt-person/mgt-person';
import { PersonCardInteraction } from '../mgt-person/mgt-person';
import { MgtTemplatedComponent } from '../templatedComponent';
import { styles } from './mgt-people-css';

/**
 * web component to display a group of people or contacts by using their photos or initials.
 *
 * @export
 * @class MgtPeople
 * @extends {MgtTemplatedComponent}
 */
@customElement('mgt-people')
export class MgtPeople extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * containing array of people used in the component.
   * @type {Array<MgtPersonDetails>}
   */
  @property({
    attribute: 'people',
    type: Object
  })
  public people: Array<MicrosoftGraph.User | MicrosoftGraph.Person | MicrosoftGraph.Contact> = null;

  /**
   * developer determined max people shown in component
   * @type {number}
   */
  @property({
    attribute: 'show-max',
    type: Number
  })
  public showMax: number = 3;
  private _firstUpdated = false;

  /**
   * Invoked when the element is first updated. Implement to perform one time
   * work on the element after update.
   *
   * Setting properties inside this method will trigger the element to update
   * again after this update cycle completes.
   *
   * * @param _changedProperties Map of changed properties with old values
   */
  protected firstUpdated() {
    this._firstUpdated = true;
    Providers.onProviderUpdated(() => this.loadPeople());
    this.loadPeople();
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */

  protected render() {
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
      <mgt-person
        .personDetails=${person}
        .personImage=${'@'}
        .personCardInteraction=${PersonCardInteraction.hover}
      ></mgt-person>
    `;
  }
}
