/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property } from 'lit-element';
import { repeat } from 'lit-html/directives/repeat';
import { getPeople, getPeopleFromGroup } from '../../graph/graph.people';
import { getUsersForUserIds } from '../../graph/graph.user';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import '../../styles/fabric-icon-font';
import '../mgt-person/mgt-person';
import { MgtTemplatedComponent } from '../templatedComponent';
import { PersonCardInteraction } from './../PersonCardInteraction';
import { styles } from './mgt-people-css';

/**
 * web component to display a group of people or contacts by using their photos or initials.
 *
 * @export
 * @class MgtPeople
 * @extends {MgtTemplatedComponent}
 *
 * @cssprop --list-margin - {String} List margin for component
 * @cssprop --avatar-margin - {String} Margin for each person
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
  public people: Array<MicrosoftGraph.User | MicrosoftGraph.Person | MicrosoftGraph.Contact>;

  /**
   * developer determined max people shown in component
   * @type {number}
   */
  @property({
    attribute: 'show-max',
    type: Number
  })
  public showMax: number;

  /**
   * determines if agenda events come from specific group
   * @type {string}
   */
  @property({
    attribute: 'group-id',
    type: String
  })
  public groupId: string;

  /**
   * user id array
   *
   * @memberof MgtPeople
   */
  @property({
    attribute: 'user-ids',
    converter: (value, type) => {
      return value.split(',');
    }
  })
  public get userIds(): string[] {
    return this.privateUserIds;
  }
  public set userIds(value: string[]) {
    const oldValue = this.userIds;
    this.privateUserIds = value;
    this.updateUserIds(value);
    this.requestUpdate('userIds', oldValue);
  }

  /**
   * Sets how the person-card is invoked
   * Set to PersonCardInteraction.none to not show the card
   *
   * @type {PersonCardInteraction}
   * @memberof MgtPerson
   */
  @property({
    attribute: 'person-card',
    converter: (value, type) => {
      value = value.toLowerCase();
      if (typeof PersonCardInteraction[value] === 'undefined') {
        return PersonCardInteraction.hover;
      } else {
        return PersonCardInteraction[value];
      }
    }
  })
  public personCardInteraction: PersonCardInteraction = PersonCardInteraction.hover;

  constructor() {
    super();

    this.showMax = 3;
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */

  protected render() {
    if (this.people && this.people.length) {
      return (
        this.renderTemplate('default', { people: this.people }) ||
        html`
          <ul class="people-list">
            ${repeat(
              this.people.slice(0, this.showMax),
              p => p.id,
              p => html`
                <li class="people-person">
                  ${this.renderTemplate('person', { person: p }, p.id) || this.renderPerson(p)}
                </li>
              `
            )}
            ${this.people.length > this.showMax
              ? this.renderTemplate('overflow', {
                  extra: this.people.length - this.showMax,
                  max: this.showMax,
                  people: this.people
                }) ||
                html`
                    <li class="overflow"><span>+${this.people.length - this.showMax}<span></li>
                  `
              : null}
          </ul>
        `
      );
    } else {
      return this.renderTemplate('no-data', null) || html``;
    }
  }

  /**
   * load state into the component.
   *
   * @protected
   * @returns
   * @memberof MgtPeople
   */
  protected async loadState() {
    if (!this.people) {
      const provider = Providers.globalProvider;

      if (provider && provider.state === ProviderState.SignedIn) {
        const graph = provider.graph.forComponent(this);

        if (this.groupId) {
          this.people = await getPeopleFromGroup(graph, this.groupId);
        } else if (this.userIds) {
          this.people = await getUsersForUserIds(graph, this.userIds);
        } else {
          this.people = await getPeople(graph);
        }
      }
    }

    this.hasLoaded = true;
  }

  private async updateUserIds(newIds: string[]) {
    if (!this.hasLoaded) {
      return;
    }

    const newIdsSet = new Set(newIds);
    this.people = this.people.filter(p => newIdsSet.has(p.id));
    const oldIdsSet = new Set(this.people ? this.people.map(p => p.id) : []);

    const newToLoad = [];

    for (const id of newIds) {
      if (!oldIdsSet.has(id)) {
        newToLoad.push(id);
      }
    }

    if (newToLoad && newToLoad.length > 0) {
      const provider = Providers.globalProvider;
      if (!provider || provider.state !== ProviderState.SignedIn) {
        return;
      }

      const graph = provider.graph.forComponent(this);

      const newPeople = await getUsersForUserIds(graph, newToLoad);
      this.people = (this.people || []).concat(newPeople);
    }
  }

  private renderPerson(person: MicrosoftGraph.Person) {
    // set image to @ to flag the mgt-person component to
    // query the image from the graph
    return html`
      <mgt-person
        .personDetails=${person}
        .personImage=${'@'}
        .personCardInteraction=${this.personCardInteraction}
      ></mgt-person>
    `;
  }
}
