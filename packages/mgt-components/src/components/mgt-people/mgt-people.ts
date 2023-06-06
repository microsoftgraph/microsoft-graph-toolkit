/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { html, TemplateResult } from 'lit';
import { property, state } from 'lit/decorators.js';
import { repeat } from 'lit/directives/repeat.js';
import { getPeople, getPeopleFromResource, PersonType } from '../../graph/graph.people';
import { getUsersPresenceByPeople } from '../../graph/graph.presence';
import { findGroupMembers, getUsersForPeopleQueries, getUsersForUserIds } from '../../graph/graph.user';
import { IDynamicPerson } from '../../graph/types';
import {
  Providers,
  ProviderState,
  MgtTemplatedComponent,
  arraysAreEqual,
  mgtHtml,
  customElement
} from '@microsoft/mgt-element';
import '../../styles/style-helper';
import { PersonCardInteraction } from './../PersonCardInteraction';
import { styles } from './mgt-people-css';
import { MgtPerson } from '../mgt-person/mgt-person';

export { PersonCardInteraction } from './../PersonCardInteraction';

/**
 * web component to display a group of people or contacts by using their photos or initials.
 *
 * @export
 * @class MgtPeople
 * @extends {MgtTemplatedComponent}
 *
 * @cssprop --people-list-margin- {String} the margin around the list of people. Default is 8px 4px 8px 8px.
 * @cssprop --people-avatar-gap - {String} the gap between the people in the list. Default is 4px.
 * @cssprop --people-overflow-font-color - {Color} the color of the overflow text.
 * @cssprop --people-overflow-font-size - {String} the text color of the overflow text. Default is 12px.
 * @cssprop --people-overflow-font-weight - {String} the font weight of the overflow text. Default is 400.
 */
@customElement('people')
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
   *
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
    void this.requestStateUpdate(true);
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
    void this.requestStateUpdate(true);
  }

  /**
   * containing array of people used in the component.
   *
   * @type {IDynamicPerson[]}
   */
  @property({
    attribute: 'people',
    type: Object
  })
  public people: IDynamicPerson[];

  /**
   * allows developer to define queries of people for component
   *
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
    void this.requestStateUpdate(true);
  }

  /**
   * developer determined max people shown in component
   *
   * @type {number}
   */
  @property({
    attribute: 'show-max',
    type: Number
  })
  public showMax: number;

  /**
   * determines if person component renders presence
   *
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
        return PersonCardInteraction[value] as PersonCardInteraction;
      }
    }
  })
  public personCardInteraction: PersonCardInteraction = PersonCardInteraction.hover;

  /**
   * The resource to get
   *
   * @type {string}
   * @memberof MgtPeople
   */
  @property({
    attribute: 'resource',
    type: String
  })
  public get resource(): string {
    return this._resource;
  }
  public set resource(value) {
    if (this._resource === value) {
      return;
    }
    this._resource = value;
    void this.requestStateUpdate(true);
  }

  /**
   * Api version to use for request
   *
   * @type {string}
   * @memberof MgtPeople
   */
  @property({
    attribute: 'version',
    type: String
  })
  public get version(): string {
    return this._version;
  }
  public set version(value) {
    if (this._version === value) {
      return;
    }
    this._version = value;
    void this.requestStateUpdate(true);
  }

  /**
   * The scopes to request
   *
   * @type {string[]}
   * @memberof MgtPeople
   */
  @property({
    attribute: 'scopes',
    converter: value => {
      return value ? value.toLowerCase().split(',') : null;
    },
    reflect: true
  })
  public scopes: string[] = [];

  /**
   * Fallback when no user is found
   *
   * @type {IDynamicPerson[]}
   */
  @property({
    attribute: 'fallback-details',
    type: Array
  })
  public get fallbackDetails(): IDynamicPerson[] {
    return this._fallbackDetails;
  }
  public set fallbackDetails(value: IDynamicPerson[]) {
    if (value === this._fallbackDetails) {
      return;
    }

    this._fallbackDetails = value;

    void this.requestStateUpdate();
  }

  /**
   * Get the scopes required for people
   *
   * @static
   * @return {*}  {string[]}
   * @memberof MgtPeople
   */
  public static get requiredScopes(): string[] {
    return [
      ...new Set([
        'user.read.all',
        'people.read',
        'user.readbasic.all',
        'presence.read.all',
        'contacts.read',
        ...MgtPerson.requiredScopes
      ])
    ];
  }

  private _groupId: string;
  private _userIds: string[];
  private _peopleQueries: string[];
  private _peoplePresence: Record<string, MicrosoftGraph.Presence> = {};
  private _resource: string;
  private _version = 'v1.0';
  private _fallbackDetails: IDynamicPerson[];
  @state() private _arrowKeyLocation = -1;

  constructor() {
    super();
    this.showMax = 3;
  }

  /**
   * Clears the state of the component
   *
   * @protected
   * @memberof MgtPeople
   */
  protected clearState(): void {
    this.people = null;
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
    const maxPeople = this.people.slice(0, this.showMax).filter(pple => pple);
    return html`
      <ul
        tabindex="0"
        class="people-list"
        aria-label="people"
        @keydown=${this.handleKeyDown}>
        ${repeat(
          maxPeople,
          p => (p.id ? p.id : p.displayName),
          p => html`
            <li tabindex="-1" class="people-person">
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
        <li tabindex="-1" aria-label="and ${extra} more attendees" class="overflow"><span>+${extra}</span></li>
      `
    );
  }

  /**
   * Handles the keypresses on a keyboard for the listed people.
   *
   * @param event is a KeyboardEvent.
   */
  protected handleKeyDown = (event: KeyboardEvent) => {
    const peopleContainer: HTMLElement = this.shadowRoot.querySelector('.people-list');
    let person: HTMLElement;
    const peopleElements: HTMLCollection = peopleContainer?.children;
    // Default all tabindex values in li nodes to -1
    for (const element of peopleElements) {
      const el: HTMLElement = element as HTMLElement;
      el.setAttribute('tabindex', '-1');
      el.blur();
    }

    const childElementCount = peopleContainer.childElementCount;
    const keyName = event.key;
    if (keyName === 'ArrowRight') {
      this._arrowKeyLocation = (this._arrowKeyLocation + 1 + childElementCount) % childElementCount;
    } else if (keyName === 'ArrowLeft') {
      this._arrowKeyLocation = (this._arrowKeyLocation - 1 + childElementCount) % childElementCount;
    } else if (keyName === 'Tab' || keyName === 'Escape') {
      this._arrowKeyLocation = -1;
      peopleContainer.blur();
    } else if (['Enter', 'space', ' '].includes(keyName)) {
      if (this.personCardInteraction !== PersonCardInteraction.none) {
        const personEl = peopleElements[this._arrowKeyLocation] as HTMLElement;
        const mgtPerson = personEl.querySelector<MgtPerson>('mgt-person');
        if (mgtPerson) {
          mgtPerson.showPersonCard();
        }
      }
    }

    if (this._arrowKeyLocation > -1) {
      person = peopleElements[this._arrowKeyLocation] as HTMLElement;
      person.setAttribute('tabindex', '1');
      person.focus();
    }
  };

  /**
   * Render an individual person.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPeople
   */
  protected renderPerson(person: MicrosoftGraph.User | MicrosoftGraph.Person | MicrosoftGraph.Contact): TemplateResult {
    let personPresence: MicrosoftGraph.Presence = {
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
      mgtHtml`
        <mgt-person
          .personDetails=${person}
          .fetchImage=${true}
          .avatarSize=${avatarSize}
          .personCardInteraction=${this.personCardInteraction}
          .showPresence=${this.showPresence}
          .personPresence=${personPresence}
          .usage=${'people'}
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
          this.people = await findGroupMembers(graph, null, this.groupId, this.showMax, PersonType.person);
        } else if (this.userIds || this.peopleQueries) {
          this.people = this.userIds
            ? await getUsersForUserIds(graph, this.userIds, '', '', this._fallbackDetails)
            : await getUsersForPeopleQueries(graph, this.peopleQueries, this._fallbackDetails);
        } else if (this.resource) {
          this.people = await getPeopleFromResource(graph, this.version, this.resource, this.scopes);
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
