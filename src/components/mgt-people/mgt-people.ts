/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property, TemplateResult } from 'lit-element';
import { repeat } from 'lit-html/directives/repeat';
import { getPeople, getPeopleFromGroup } from '../../graph/graph.people';
import { getUsersForUserIds } from '../../graph/graph.user';
import { MgtTemplatedComponent, Providers, ProviderState } from '../../mgt-core';
import '../../styles/fabric-icon-font';
import '../mgt-person/mgt-person';
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

  private privateUserIds: string[];

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
    if (this.isLoadingState) {
      return this.renderLoading();
    }

    if (!this.people || this.people.length === 0) {
      return this.renderNoData();
    }

    return this.renderTemplate('default', { people: this.people, max: this.showMax }) || this.renderPeople();
  }

  /**
   * Render the loading state.
   *
   * @protected
   * @returns
   * @memberof MgtPeople
   */
  protected renderLoading() {
    return this.renderTemplate('loading', null) || html``;
  }

  /**
   * Render the list of people.
   *
   * @protected
   * @param {*} people
   * @returns {TemplateResult}
   * @memberof MgtPeople
   */
  protected renderPeople(): TemplateResult {
    const maxPeople = this.people.slice(0, this.showMax);

    return html`
      <ul class="people-list">
        ${repeat(
          maxPeople,
          p => p.id,
          p => html`
            <li class="people-person">
              ${this.renderPerson(p)}
            </li>
          `
        )}
        ${this.people.length > this.showMax ? this.renderOverflow() : null}
      </ul>
    `;
  }

  /**
   * Render the overflow content to represent any extra people, beyond the max.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPeople
   */
  protected renderOverflow(): TemplateResult {
    const extra = this.people.length - this.showMax;
    return (
      this.renderTemplate('overflow', {
        extra,
        max: this.showMax,
        people: this.people
      }) ||
      html`
        <li class="overflow"><span>+${extra}<span></li>
      `
    );
  }

  /**
   * Render an individual person.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPeople
   */
  protected renderPerson(person: MicrosoftGraph.User | MicrosoftGraph.Person | MicrosoftGraph.Contact): TemplateResult {
    return (
      this.renderTemplate('person', { person }, person.id) ||
      // set image to @ to flag the mgt-person component to
      // query the image from the graph
      html`
        <mgt-person
          .personDetails=${person}
          .personImage=${'@'}
          .personCardInteraction=${this.personCardInteraction}
        ></mgt-person>
      `
    );
  }

  /**
   * render the no data state.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPeople
   */
  protected renderNoData(): TemplateResult {
    return this.renderTemplate('no-data', null) || html``;
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
  }

  private async updateUserIds(newIds: string[]) {
    if (this.isLoadingState) {
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
}
