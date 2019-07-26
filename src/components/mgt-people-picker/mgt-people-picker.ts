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
import { styles } from './mgt-people-picker-css';

import '../mgt-person/mgt-person';
import '../../styles/fabric-icon-font';
import { MgtTemplatedComponent } from '../templatedComponent';
import { MgtPersonDetails, MgtPerson } from '../mgt-person/mgt-person';

@customElement('mgt-people-picker')
export class MgtPicker extends MgtTemplatedComponent {
  @property({
    attribute: 'people',
    type: Object
  })
  people: Array<MgtPersonDetails> = null;

  @property({
    attribute: 'show-max',
    type: Number
  })
  showMax: number = 6;

  private arrowSelectionCount: number = 0;

  @property({
    attribute: 'group',
    type: String
  })
  group: string;

  @property() private _personName: string = '';
  @property() private _selectedPeople: Array<any> = [];
  @property() private _duplicatePersonId: string = '';
  @property() private _userInput: string = '';
  @property() private _previousSearch: any;
  @property() private isUserFocused: boolean = false;

  static get styles() {
    return styles;
  }

  constructor() {
    super();
    this.trackMouseFocus = this.trackMouseFocus.bind(this);
  }

  private onUserTypeSearch(event: any) {
    if (event.keyCode == 40 || event.keyCode == 38) {
      this.handleArrowSelection(event);
      return;
    }
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
      let newEvent = new CustomEvent('selectionChanged', {
        bubbles: true,
        cancelable: false,
        detail: this._selectedPeople
      });
      this.dispatchEvent(newEvent);
      return;
    }
    this._userInput = event.target.value;
    if (event.target.value) {
      this.loadPersonSearch(this._userInput);
    } else {
      event.target.value = '';
      this._userInput = '';
      this.people = [];
    }
  }

  private onUserKeyDown(event: any) {
    if (event.code == 'Tab') {
      this.addPerson(this.people[this.arrowSelectionCount], event);
      event.target.value = '';
      event.preventDefault();
    }
  }

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
        if (this.arrowSelectionCount <= this.people.length - this.showMax) {
          this.arrowSelectionCount++;
        } else {
          this.arrowSelectionCount = 0;
        }
      }

      const peopleList = this.renderRoot.querySelector('.people-list');
      //reset background color
      for (let i = 0; i < peopleList.children.length; i++) {
        peopleList.children[i].setAttribute('style', 'background-color: transparent ');
      }
      //set selected background
      peopleList.children[this.arrowSelectionCount].setAttribute('style', 'background-color: #f1f1f1;');
    }
  }

  private addPerson(person: MgtPersonDetails, event: any) {
    if (person) {
      this._personName = '';
      this._duplicatePersonId = '';
      let chosenPerson: any = person;
      let filteredPersonArr = this._selectedPeople.filter(function(person) {
        return person.id == chosenPerson.id;
      });
      if (this._selectedPeople.length && filteredPersonArr.length) {
        this._duplicatePersonId = chosenPerson.id;
      } else {
        this._selectedPeople.push(person);
        let event = new CustomEvent('selectionChanged', {
          bubbles: true,
          cancelable: false,
          detail: this._selectedPeople
        });
        this.dispatchEvent(event);

        this.people = [];
        this._userInput = '';
        this._personName = '';
        this.arrowSelectionCount = 0;
      }
    }
  }

  private async loadPersonSearch(name: string) {
    let provider = Providers.globalProvider;
    let peoples: any;
    if (provider && provider.state === ProviderState.SignedIn) {
      let client = Providers.globalProvider.graph;
      //determine if group property is requested

      if (this.group) {
        peoples = await client.getPeopleFromGroup(this.group).catch(function() {
          return;
        });
      } else {
        peoples = await client.findPerson(name).catch(function() {
          return;
        });
      }
      if (peoples.length) {
        this.filterPeople(peoples);
      } else {
        this.people = [];
      }
    }
  }

  private filterPeople(peoples: any) {
    //check if people need to be updated
    if (this.people) {
      if (this.people.length > 0) {
        this._previousSearch = this.people;
      } else {
        this._previousSearch = [''];
      }
      //find ids from selected people
      let id_filter = this._selectedPeople.map(function(el) {
        return el.id;
      });
      //filter id's
      let filtered = peoples.filter(function(person) {
        return id_filter.indexOf(person.id) === -1;
      });
      if (filtered.length == 0 && this._userInput.length > 0) {
        this.people = [];
        return;
      } else {
        if (filtered.length) {
          filtered[0].isSelected = 'fill';
          this.arrowSelectionCount = 0;
          this.people = filtered;
        }
      }
    } else {
      peoples[0].isSelected = 'fill';
      this.arrowSelectionCount = 0;
      this.people = peoples;
    }
  }

  private removePerson(person: MgtPersonDetails) {
    let chosenPerson: any = person;
    let filteredPersonArr = this._selectedPeople.filter(function(person) {
      return person.id !== chosenPerson.id;
    });
    this._selectedPeople = filteredPersonArr;
    let event = new CustomEvent('selectionChanged', {
      bubbles: true,
      cancelable: false,
      detail: this._selectedPeople
    });
    this.dispatchEvent(event);
    this.renderChosenPeople();
  }

  private renderErrorMessage() {
    if (this.people) {
      if (this.people.length === 0 && this._userInput.length > 0) {
        return html`
          <div class="error-message-parent">
            <div class="search-error-text">We didn't find any matches.</div>
          </div>
        `;
      }
    }
  }
  private trackMouseFocus(e) {
    if (e.target.localName === 'mgt-people-picker') {
      //Mouse is focused on input
      this.isUserFocused = true;
      if (this._userInput) {
        this.loadPersonSearch(this._userInput);
      }
    } else {
      //reset if not clicked in focus
      this.isUserFocused = false;
      this.people = [];
    }
  }

  private renderChosenPeople() {
    let peopleList;
    if (this._selectedPeople.length > 0) {
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
      <ul class="people-chosen-list">
        ${peopleList}
        <div class="input-search-start">
          <input
            id="people-picker-input"
            class="people-chosen-input"
            type="text"
            placeholder="Start typing a name"
            .value="${this._personName}"
            @keydown="${(e: KeyboardEvent & { target: HTMLInputElement }) => {
              this.onUserKeyDown(e);
            }}"
            @keyup="${(e: KeyboardEvent & { target: HTMLInputElement }) => {
              this.onUserTypeSearch(e);
            }}"
          />
        </div>
      </ul>
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
    } else {
      if (/\s/g.test(this._userInput) == true) {
        //if highlight is not found due to space character
        let search = this._userInput.replace(/\s/g, '');
        this._userInput = search;
      } else {
        peoples.first = peoples.displayName;
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
    let peoples: any = this.people;
    if (peoples) {
      return html`
        <ul class="people-list">
          ${this.renderPersons(peoples)}
        </ul>
      `;
    }
  }
  private renderPersons(peoples: any) {
    return peoples.slice(0, this.showMax).map(
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
    document.addEventListener('mousedown', this.trackMouseFocus, false);
    return (
      this.renderTemplate('default', { people: this.people }) ||
      html`
        <div class="people-picker">
          <div class="people-picker-input">
            ${this.renderChosenPeople()}
          </div>
          <div class="people-list-separator"></div>
          ${this.isUserFocused ? this.renderPeopleList() : null}
          <div class="error-message-holder">
            ${this._userInput.length !== 0 && this.isUserFocused ? this.renderErrorMessage() : null}
          </div>
        </div>
      `
    );
  }

  private renderPerson(person: MicrosoftGraph.Person) {
    return html`
      <mgt-person person-details=${JSON.stringify(person)}></mgt-person>
    `;
  }
  private renderChosenPerson(person: MicrosoftGraph.Person) {
    return html`
      <mgt-person class="chosen-person" person-details=${JSON.stringify(person)}></mgt-person>
    `;
  }
}
