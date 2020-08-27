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
import { getUsersPresenceByPeople } from '../../graph/graph.presence';
import { getUsersForPeopleQueries, getUsersForUserIds } from '../../graph/graph.user';
import { IDynamicPerson } from '../../graph/types';
import { Providers, ProviderState } from '@microsoft/mgt-element';
import '../../styles/fabric-icon-font';
import { arraysAreEqual } from '../../utils/Utils';
import { MgtTemplatedComponent } from '../templatedComponent';
import { PersonCardInteraction } from './../PersonCardInteraction';
import { styles } from './mgt-people-css';

export { PersonCardInteraction } from './../PersonCardInteraction';

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
   * determines if agenda events come from specific group
   * @type {string}
   */
  @property({
    attribute: 'group-id',
    type: String
  })
  public get groupId(): string {
    return this._groupId;
  }
  public set groupId(value) {
    if (this._groupId === value) {
      return;
    }
    this._groupId = value;
    this.requestStateUpdate(true);
  }

  /**
   * user id array
   *
   * @memberof MgtPeople
   */
  @property({
    attribute: 'user-ids',
    converter: (value, type) => {
      return value.split(',').map(v => v.trim());
    }
  })
  public get userIds(): string[] {
    return this._userIds;
  }
  public set userIds(value: string[]) {
    if (arraysAreEqual(this._userIds, value)) {
      return;
    }
    this._userIds = value;
    this.requestStateUpdate(true);
  }

  /**
   * containing array of people used in the component.
   * @type {IDynamicPerson[]}
   */
  @property({
    attribute: 'people',
    type: Object
  })
  public people: IDynamicPerson[];

  /**
   * allows developer to define queries of people for component
   * @type {string[]}
   */

  @property({
    attribute: 'people-queries',
    converter: (value, type) => {
      return value.split(',').map(v => v.trim());
    }
  })
  public get peopleQueries(): string[] {
    return this._peopleQueries;
  }
  public set peopleQueries(value: string[]) {
    if (arraysAreEqual(this._peopleQueries, value)) {
      return;
    }
    this._peopleQueries = value;
    this.requestStateUpdate(true);
  }

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
   * determines if person component renders presence
   * @type {boolean}
   */
  @property({
    attribute: 'show-presence',
    type: Boolean
  })
  public showPresence: boolean;

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

  private _groupId: string;
  private _userIds: string[];
  private _peopleQueries: string[];
  private _peoplePresence: {};

  constructor() {
    super();

    this.showMax = 3;
  }

  /**
   * Request to reload the state.
   * Use reload instead of load to ensure loading events are fired.
   *
   * @protected
   * @memberof MgtBaseComponent
   */
  protected requestStateUpdate(force?: boolean) {
    if (force) {
      this.people = null;
    }
    return super.requestStateUpdate(force);
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
    let personPresence = {
      // set up default presence
      activity: 'Offline',
      availability: 'Offline',
      id: null
    };
    if (this.showPresence && this._peoplePresence) {
      personPresence = this._peoplePresence[person.id];
    }
    const avatarSize = 'small';
    return (
      this.renderTemplate('person', { person }, person.id) ||
      // set image to @ to flag the mgt-person component to
      // query the image from the graph
      html`
        <mgt-person
          .personDetails=${person}
          .fetchImage=${true}
          .avatarSize=${avatarSize}
          .personCardInteraction=${this.personCardInteraction}
          .showPresence=${this.showPresence}
          .personPresence=${personPresence}
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

        // populate people
        if (this.groupId) {
          this.people = await getPeopleFromGroup(graph, this.groupId);
        } else if (this.userIds) {
          this.people = await getUsersForUserIds(graph, this.userIds);
        } else if (this.peopleQueries) {
          this.people = await getUsersForPeopleQueries(graph, this.peopleQueries);
        } else {
          this.people = await getPeople(graph);
        }

        // populate presence for people
        if (this.showPresence) {
          this._peoplePresence = await getUsersPresenceByPeople(graph, this.people);
        } else {
          this._peoplePresence = null;
        }
      }
    }
  }
}
