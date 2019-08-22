/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html, customElement, property } from 'lit-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import { styles } from './mgt-people-picker-css';

import '../mgt-person/mgt-person';
import '../../styles/fabric-icon-font';
import { MgtTemplatedComponent } from '../templatedComponent';
import { MgtPersonDetails, MgtPerson } from '../mgt-person/mgt-person';
import { debounce } from '../../utils/utils';

@customElement('mgt-people-picker')
export class MgtPicker extends MgtTemplatedComponent {
  /**
   * people property, containing object of MgtPersonDetails.
   * @type {MgtPersonDetails}
   */
  @property({
    attribute: 'people',
    type: Object
  })
  people: Array<MgtPersonDetails> = null;

  /**
   * show-max property, determining how many people to show in list.
   * @type {number}
   */
  @property({
    attribute: 'show-max',
    type: Number
  })
  showMax: number = 6;

  /**
   * group-id property, value determining if search is filtered to a group.
   * @type {string}
   */
  @property({
    attribute: 'group-id',
    type: String
  })
  groupId: string;

  /**
   * _selectedPeople property, array of user picked people.
   * @type {Array<any>}
   */
  @property() public _selectedPeople: Array<any> = [];

  /**
   * _duplicatePersonId property, value checking if user has already selected person's id.
   * @type {string}
   */
  @property() private _duplicatePersonId: string = '';

  /**
   * _userInput property, containing the current input value of search box.
   * @type {string}
   */
  @property() private _userInput: string = '';

  /**
   * arrowSelectionCount property, current highlighted value of person in search area.
   * @type {number}
   */
  private arrowSelectionCount: number = 0;

  /**
   * groupPeople property, contains people in specific group
   * @type {any[]}
   */
  private groupPeople: any[];

  /**
   * isLoading property, determines if search is still loading people details
   * @type {boolean}
   */
  private isLoading: boolean = false;

  /**
   * Adds debounce method for set delay on user input
   * @returns userinput
   */
  private debounceHandle = window.addEventListener(
    'keyup',
    debounce(() => {
      this.loadPersonSearch(this._userInput);
    }, 300)
  );

  attributeChangedCallback(att, oldval, newval) {
    super.attributeChangedCallback(att, oldval, newval);

    if (att == 'group-id' && oldval !== newval) {
      this.findGroup();
    }
  }

  static get styles() {
    return styles;
  }

  constructor() {
    super();
    this.trackMouseFocus = this.trackMouseFocus.bind(this);
  }

  firstUpdated() {
    if (this.groupId) {
      Providers.onProviderUpdated(() => this.findGroup());
      this.findGroup();
    }
  }
  /**
   * Async query to Graph for members of group if determined by developer.
   * set's `this.groupPeople` to those members.
   */
  private async findGroup() {
    let provider = Providers.globalProvider;
    if (provider && provider.state === ProviderState.SignedIn) {
      let client = Providers.globalProvider.graph;
      this.groupPeople = await client.getPeopleFromGroup(this.groupId);
    }
  }
  /**
   * Tracks event on user input in search
   * @param event - event tracked when user input is detected (keyup)
   */
  private onUserTypeSearch(event: any) {
    if (event.code == 'Escape') {
      event.target.value = '';
      this._userInput = '';
      this.people = [];
      return;
    }
    if (event.code == 'Backspace' && this._userInput.length == 0 && this._selectedPeople.length > 0) {
      event.target.value = '';
      this._userInput = '';
      //remove last person in selected list
      this._selectedPeople = this._selectedPeople.splice(0, this._selectedPeople.length - 1);
      //fire selected people changed event
      this.fireCustomEvent('selectionChanged', this._selectedPeople);
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
    if (event.keyCode == 40 || event.keyCode == 38) {
      //keyCodes capture: down arrow (40) and up arrow (38)
      this.handleArrowSelection(event);
      if (this._userInput.length > 0) {
        event.preventDefault();
      }
    }
    if (event.code == 'Tab' || event.code == 'Enter') {
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
      //update arrow count
      if (event.keyCode == 38) {
        //up arrow
        if (this.arrowSelectionCount > 0) {
          this.arrowSelectionCount--;
        } else {
          this.arrowSelectionCount = 0;
        }
      }
      if (event.keyCode == 40) {
        //down arrow
        if (this.arrowSelectionCount + 1 !== this.people.length && this.arrowSelectionCount + 1 < this.showMax) {
          this.arrowSelectionCount++;
        } else {
          this.arrowSelectionCount = 0;
        }
      }

      const peopleList = this.renderRoot.querySelector('.people-list');
      //reset background color
      for (let i = 0; i < peopleList.children.length; i++) {
        peopleList.children[i].setAttribute('class', 'people-person-list');
      }
      //set selected background
      peopleList.children[this.arrowSelectionCount].setAttribute('class', 'people-person-list-fill');
    }
  }

  /**
   * Tracks when user selects person from picker
   * @param person - contains details pertaining to selected user
   * @param event - tracks user event
   */
  private addPerson(person: MgtPersonDetails, event: any) {
    if (person) {
      this._userInput = '';
      this._duplicatePersonId = '';
      let chosenPerson: any = person;
      let filteredPersonArr = this._selectedPeople.filter(function(person) {
        return person.id == chosenPerson.id;
      });
      if (this._selectedPeople.length && filteredPersonArr.length) {
        this._duplicatePersonId = chosenPerson.id;
      } else {
        this._selectedPeople.push(person);
        this.fireCustomEvent('selectionChanged', this._selectedPeople);

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
      let provider = Providers.globalProvider;
      let people: any;

      if (provider && provider.state === ProviderState.SignedIn) {
        this.isLoading = true;
        let client = Providers.globalProvider.graph;

        //filtering groups
        if (this.groupId) {
          people = this.groupPeople;
        } else {
          people = await client.findPerson(name);
        }

        if (people) {
          people = people.filter(function(person) {
            return person.displayName.toLowerCase().indexOf(name) !== -1;
          });
        }

        for (let person of people) {
          // set image to @ to flag the mgt-person component to
          // query the image from the graph
          person.image = '@';
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
    //check if people need to be updated
    //ensuring people list is displayed
    //find ids from selected people
    if (people) {
      let id_filter = this._selectedPeople.map(function(el) {
        return el.id;
      });

      //filter id's
      let filtered = people.filter(function(person) {
        return id_filter.indexOf(person.id) === -1;
      });

      return filtered;
    }
  }

  /**
   * Removes person from selected people
   * @param person - person and details pertaining to user selected
   */
  private removePerson(person: MgtPersonDetails) {
    let chosenPerson: any = person;
    let filteredPersonArr = this._selectedPeople.filter(function(person) {
      return person.id !== chosenPerson.id;
    });
    this._selectedPeople = filteredPersonArr;
    this.fireCustomEvent('selectionChanged', this._selectedPeople);
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
  private trackMouseFocus(e) {
    const peopleList = this.renderRoot.querySelector('.people-list');
    if (e.target.localName === 'mgt-people-picker') {
      const peopleInput = <HTMLInputElement>this.renderRoot.querySelector('.people-chosen-input');
      peopleInput.focus();
      peopleInput.select();
    }
    if (peopleList) {
      if (e.target.localName === 'mgt-people-picker') {
        //Mouse is focused on input
        peopleList.setAttribute('style', 'display:block');
      } else {
        //reset if not clicked in focus
        peopleList.setAttribute('style', 'display:none');
      }
    }
  }

  private renderChosenPeople() {
    let peopleList;
    let inputClass = 'input-search-start';
    if (this._selectedPeople.length > 0) {
      inputClass = 'input-search';
      peopleList = html`
        ${this._selectedPeople.slice(0, this._selectedPeople.length).map(
          person =>
            html`
              <li class="${person.id == this._duplicatePersonId ? 'people-person duplicate-person' : 'people-person'}">
                ${this.renderTemplate('person', { person: person }, person.displayName) ||
                  this.renderChosenPerson(person)}
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

  private renderHighlightText(person: MgtPersonDetails) {
    let peoples: any = person;

    let highlightLocation = peoples.displayName.toLowerCase().indexOf(this._userInput.toLowerCase());
    if (highlightLocation !== -1) {
      //no location
      if (highlightLocation == 0) {
        //highlight is at the beginning of sentence
        peoples.first = '';
        peoples.highlight = peoples.displayName.slice(0, this._userInput.length);
        peoples.last = peoples.displayName.slice(this._userInput.length, peoples.displayName.length);
      } else if (highlightLocation == peoples.displayName.length) {
        //highlight is at end of the sentence
        peoples.first = peoples.displayName.slice(0, highlightLocation);
        peoples.highlight = peoples.displayName.slice(highlightLocation, peoples.displayName.length);
        peoples.last = '';
      } else {
        //highlight is in middle of sentence
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
    let people = this.people;
    if (people) {
      if (people.length == 0 && this._userInput.length > 0 && this.isLoading == false) {
        return html`
          <div class="people-list">
            ${this.renderErrorMessage()}
          </div>
        `;
      } else {
        return html`
          <div class="people-list">
            ${this.renderPersons(people)}
          </div>
        `;
      }
    }
  }

  private renderPersons(people: Array<any>) {
    return people.slice(0, this.showMax).map(
      person =>
        html`
          <li
            class="${person.isSelected == 'fill' ? 'people-person-list-fill' : 'people-person-list'}"
            @click="${(event: any) => this.addPerson(person, event)}"
          >
            ${this.renderTemplate('person', { person: person }, person.displayName) || this.renderPerson(person)}
            <div class="people-person-text-area" id="${person.displayName}">
              ${this.renderHighlightText(person)}
              <span class="people-person-job-title">${person.jobTitle}</span>
            </div>
          </li>
        `
    );
  }

  render() {
    document.addEventListener('mouseup', this.trackMouseFocus, false);
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

  private renderPerson(person: MicrosoftGraph.Person) {
    return html`
      <mgt-person .personDetails=${person}></mgt-person>
    `;
  }
  private renderChosenPerson(person: MicrosoftGraph.Person) {
    return html`
      <mgt-person class="chosen-person" .personDetails=${person}></mgt-person>
    `;
  }
}
