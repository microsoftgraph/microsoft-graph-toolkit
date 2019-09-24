/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property } from 'lit-element';
import { repeat } from 'lit-html/directives/repeat';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import '../../styles/fabric-icon-font';
import { debounce } from '../../utils/utils';
import '../mgt-person/mgt-person';
import { MgtTemplatedComponent } from '../templatedComponent';
import { styles } from './mgt-people-picker-css';

/**
 * Web component used to search for people from the Microsoft Graph
 *
 * @export
 * @class MgtPicker
 * @extends {MgtTemplatedComponent}
 */
@customElement('mgt-people-picker')
export class MgtPeoplePicker extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * user the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * containing object of MgtPersonDetails.
   * @type {MgtPersonDetails}
   */
  @property({
    attribute: 'people',
    type: Object
  })
  public people: Array<MicrosoftGraph.User | MicrosoftGraph.Person | MicrosoftGraph.Contact> = null;

  /**
   * determining how many people to show in list.
   * @type {number}
   */
  @property({
    attribute: 'show-max',
    type: Number
  })
  public showMax: number = 6;

  /**
   * value determining if search is filtered to a group.
   * @type {string}
   */
  @property({
    attribute: 'group-id',
    type: String
  })
  public groupId: string;

  /**
   *  array of user picked people.
   * @type {Array<any>}
   */
  @property() public selectedPeople: Array<MicrosoftGraph.User | MicrosoftGraph.Person | MicrosoftGraph.Contact> = [];

  // single matching id for filtering against people list
  @property() private _duplicatePersonId: string = '';

  // User input in search
  @property() private _userInput: string = '';

  // tracking of user arrow key input for selection
  private arrowSelectionCount: number = 0;
  // List of people requested if group property is provided
  private groupPeople: any[];
  // if search is still loading don't load "people not found" state
  private isLoading = false;

  // handing debounce on user search

  /**
   * Adds debounce method for set delay on user input
   * @returns userinput
   */
  private debounceHandle = window.addEventListener(
    'keyup',
    debounce(e => {
      if (e.keyCode === 40 || e.keyCode === 38) {
        // keyCodes capture: down arrow (40) and up arrow (38)
        return;
      } else {
        this.arrowSelectionCount = 0;
        this.loadPersonSearch(this._userInput);
      }
    }, 300)
  );

  constructor() {
    super();
  }

  /**
   * Synchronizes property values when attributes change.
   *
   * @param {*} name
   * @param {*} oldValue
   * @param {*} newValue
   * @memberof MgtPersonCard
   */
  public attributeChangedCallback(att, oldval, newval) {
    super.attributeChangedCallback(att, oldval, newval);

    if (att === 'group-id' && oldval !== newval) {
      this.findGroup();
    }
  }

  public firstUpdated() {
    if (this.groupId) {
      Providers.onProviderUpdated(() => this.findGroup());
      this.findGroup();
    }
  }

  public render() {
    return (
      this.renderTemplate('default', { people: this.people }) ||
      html`
        <div class="people-picker">
          <div class="people-picker-input">
            ${this.renderChosenPeople()}
          </div>
          <div class="people-list-separator"></div>
          ${this.renderPeopleList()}
        </div>
      `
    );
  }

  /**
   * Async query to Graph for members of group if determined by developer.
   * set's `this.groupPeople` to those members.
   */
  private async findGroup() {
    const provider = Providers.globalProvider;
    if (provider && provider.state === ProviderState.SignedIn) {
      const client = Providers.globalProvider.graph;
      this.groupPeople = await client.getPeopleFromGroup(this.groupId);
    }
  }

  /**
   * Tracks event on user input in search
   * @param event - event tracked when user input is detected (keyup)
   */
  private onUserTypeSearch(event: any) {
    if (event.code === 'Escape') {
      event.target.value = '';
      this._userInput = '';
      this.people = [];
      return;
    }
    if (event.code === 'Backspace' && this._userInput.length === 0 && this.selectedPeople.length > 0) {
      event.target.value = '';
      this._userInput = '';
      // remove last person in selected list
      this.selectedPeople = this.selectedPeople.splice(0, this.selectedPeople.length - 1);
      // fire selected people changed event
      this.fireCustomEvent('selectionChanged', this.selectedPeople);
      return;
    }
    this._userInput = event.target.value;
    if (event.target.value) {
      this.debounceHandle;
    } else {
      event.target.value = '';
      this._userInput = '';
      this.people = [];
    }
  }

  /**
   * Tracks event on user search (keydown)
   * @param event - event tracked on user input (keydown)
   */
  private onUserKeyDown(event: any) {
    if (event.keyCode === 40 || event.keyCode === 38) {
      // keyCodes capture: down arrow (40) and up arrow (38)
      this.handleArrowSelection(event);
      if (this._userInput.length > 0) {
        event.preventDefault();
      }
    }
    if (event.code === 'Tab' || event.code === 'Enter') {
      if (this.people.length) {
        event.preventDefault();
      }
      this.addPerson(this.people[this.arrowSelectionCount], event);
      event.target.value = '';
    }
  }

  /**
   * Tracks user key selection for arrow key selection of people
   * @param event - tracks user key selection
   */
  private handleArrowSelection(event: any) {
    if (this.people.length) {
      // update arrow count
      if (event.keyCode === 38) {
        // up arrow
        if (this.arrowSelectionCount > 0) {
          this.arrowSelectionCount--;
        } else {
          this.arrowSelectionCount = 0;
        }
      }
      if (event.keyCode === 40) {
        // down arrow
        if (this.arrowSelectionCount + 1 !== this.people.length && this.arrowSelectionCount + 1 < this.showMax) {
          this.arrowSelectionCount++;
        } else {
          this.arrowSelectionCount = 0;
        }
      }

      const peopleList = this.renderRoot.querySelector('.people-list');
      // reset background color
      for (let i = 0; i < peopleList.children.length; i++) {
        peopleList.children[i].setAttribute('class', 'people-person-list');
      }
      // set selected background
      peopleList.children[this.arrowSelectionCount].setAttribute('class', 'people-person-list-fill');
    }
  }

  /**
   * Tracks when user selects person from picker
   * @param person - contains details pertaining to selected user
   * @param event - tracks user event
   */
  private addPerson(person: MicrosoftGraph.User | MicrosoftGraph.Person | MicrosoftGraph.Contact, event: any) {
    if (person) {
      this._userInput = '';
      this._duplicatePersonId = '';
      const chosenPerson: any = person;
      const filteredPersonArr = this.selectedPeople.filter(person => {
        return person.id === chosenPerson.id;
      });
      if (this.selectedPeople.length && filteredPersonArr.length) {
        this._duplicatePersonId = chosenPerson.id;
      } else {
        this.selectedPeople.push(person);
        this.fireCustomEvent('selectionChanged', this.selectedPeople);

        this.people = [];
        this._userInput = '';
        this.arrowSelectionCount = 0;
      }
    }
  }

  /**
   * Async method which query's the Graph with user input
   * @param name - user input or name of person searched
   */
  private async loadPersonSearch(name: string) {
    if (name.length) {
      name = name.toLowerCase();
      const provider = Providers.globalProvider;
      let people: any;

      if (provider && provider.state === ProviderState.SignedIn) {
        const that = this;
        setTimeout(() => {
          that.isLoading = true;
        }, 400);

        const client = Providers.globalProvider.graph;

        // filtering groups
        if (this.groupId) {
          people = this.groupPeople;
        } else {
          people = await client.findPerson(name);
        }

        if (people) {
          people = people.filter(person => {
            return person.displayName.toLowerCase().indexOf(name) !== -1;
          });
        }

        this.people = this.filterPeople(people);
        this.isLoading = false;
      }
    }
  }
  /**
   * Filters people searched from already selected people
   * @param people - array of people returned from query to Graph
   */
  private filterPeople(people: any) {
    // check if people need to be updated
    // ensuring people list is displayed
    // find ids from selected people
    if (people) {
      const idFilter = this.selectedPeople.map(el => {
        return el.id;
      });

      // filter id's
      const filtered = people.filter(person => {
        return idFilter.indexOf(person.id) === -1;
      });

      return filtered;
    }
  }

  /**
   * Removes person from selected people
   * @param person - person and details pertaining to user selected
   */
  private removePerson(person: MicrosoftGraph.User | MicrosoftGraph.Person | MicrosoftGraph.Contact) {
    const chosenPerson: any = person;
    const filteredPersonArr = this.selectedPeople.filter(person => {
      return person.id !== chosenPerson.id;
    });
    this.selectedPeople = filteredPersonArr;
    this.fireCustomEvent('selectionChanged', this.selectedPeople);
  }

  private renderErrorMessage() {
    return html`
      <div class="error-message-parent">
        <div label="search-error-text" aria-label="We didn't find any matches." class="search-error-text">
          We didn't find any matches.
        </div>
      </div>
    `;
  }

  private renderChosenPeople() {
    let peopleList;
    let inputClass = 'input-search-start';
    if (this.selectedPeople.length > 0) {
      inputClass = 'input-search';
      peopleList = html`
        ${this.selectedPeople.slice(0, this.selectedPeople.length).map(
          person =>
            html`
              <li class="${person.id === this._duplicatePersonId ? 'people-person duplicate-person' : 'people-person'}">
                ${this.renderTemplate('person', { person }, person.displayName) || this.renderChosenPerson(person)}
                <p class="person-display-name">${person.displayName}</p>
                <div class="CloseIcon" @click="${() => this.removePerson(person)}">\uE711</div>
              </li>
            `
        )}
      `;
    }
    return html`
      <div class="people-chosen-list">
        ${peopleList}
        <div class="${inputClass}">
          <input
            id="people-picker-input"
            class="people-chosen-input"
            type="text"
            placeholder="Start typing a name"
            label="people-picker-input"
            aria-label="people-picker-input"
            role="input"
            .value="${this._userInput}"
            @blur=${this.lostFocus}
            @click=${this.gainedFocus}
            @keydown="${(e: KeyboardEvent & { target: HTMLInputElement }) => {
              this.onUserKeyDown(e);
            }}"
            @keyup="${(e: KeyboardEvent & { target: HTMLInputElement }) => {
              this.onUserTypeSearch(e);
            }}"
          />
        </div>
      </div>
    `;
  }

  private gainedFocus() {
    const peopleList = this.renderRoot.querySelector('.people-list');
    const peopleInput = this.renderRoot.querySelector('.people-chosen-input') as HTMLInputElement;
    peopleInput.focus();
    peopleInput.select();
    if (peopleList) {
      // Mouse is focused on input
      peopleList.setAttribute('style', 'display:block');
    }
  }
  private lostFocus() {
    const peopleList = this.renderRoot.querySelector('.people-list');
    if (peopleList) {
      setTimeout(() => {
        peopleList.setAttribute('style', 'display:none');
      }, 300);
    }
  }

  private renderHighlightText(person: MicrosoftGraph.User | MicrosoftGraph.Person | MicrosoftGraph.Contact) {
    const peoples: any = person;

    const highlightLocation = peoples.displayName.toLowerCase().indexOf(this._userInput.toLowerCase());
    if (highlightLocation !== -1) {
      // no location
      if (highlightLocation === 0) {
        // highlight is at the beginning of sentence
        peoples.first = '';
        peoples.highlight = peoples.displayName.slice(0, this._userInput.length);
        peoples.last = peoples.displayName.slice(this._userInput.length, peoples.displayName.length);
      } else if (highlightLocation === peoples.displayName.length) {
        // highlight is at end of the sentence
        peoples.first = peoples.displayName.slice(0, highlightLocation);
        peoples.highlight = peoples.displayName.slice(highlightLocation, peoples.displayName.length);
        peoples.last = '';
      } else {
        // highlight is in middle of sentence
        peoples.first = peoples.displayName.slice(0, highlightLocation);
        peoples.highlight = peoples.displayName.slice(highlightLocation, highlightLocation + this._userInput.length);
        peoples.last = peoples.displayName.slice(
          highlightLocation + this._userInput.length,
          peoples.displayName.length
        );
      }
    }

    return html`
      <div>
        <span class="people-person-text">${peoples.first}</span
        ><span class="people-person-text highlight-search-text">${peoples.highlight}</span
        ><span class="people-person-text">${peoples.last}</span>
      </div>
    `;
  }

  private renderPeopleList() {
    let people: any = this.people;
    if (people) {
      people = people.slice(0, this.showMax);
      if (people.length === 0 && this._userInput.length > 0 && this.isLoading === false) {
        return html`
          <div class="people-list">
            ${this.renderErrorMessage()}
          </div>
        `;
      } else {
        if (people[0]) {
          (people[0] as any).isSelected = 'fill';
        }
        return html`
          <div class="people-list">
            ${this.renderPersons(people)}
          </div>
        `;
      }
    }
  }

  private renderPersons(people: any[]) {
    return html`
      ${repeat(
        people,
        person => person.id,
        person => html`
          <li
            class="${person.isSelected === 'fill' ? 'people-person-list-fill' : 'people-person-list'}"
            @click="${(event: any) => this.addPerson(person, event)}"
          >
            ${this.renderTemplate('person', { person }, person.displayName) || this.renderPerson(person)}
            <div class="people-person-text-area" id="${person.displayName}">
              ${this.renderHighlightText(person)}
              <span class="people-person-job-title">${person.jobTitle}</span>
            </div>
          </li>
        `
      )}
    `;
  }

  private renderPerson(person: MicrosoftGraph.Person) {
    return html`
      <mgt-person .personDetails=${person} .personImage=${'@'}></mgt-person>
    `;
  }
  private renderChosenPerson(person: MicrosoftGraph.Person) {
    return html`
      <mgt-person class="chosen-person" .personDetails=${person} .personImage=${'@'}></mgt-person>
    `;
  }
}
