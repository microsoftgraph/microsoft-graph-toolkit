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
import { debounce } from '../../utils/utils';

@customElement('mgt-people-picker')
export class MgtPicker extends MgtTemplatedComponent {
  //people in current search list
  @property({
    attribute: 'people',
    type: Object
  })
  people: Array<MgtPersonDetails> = null;

  //maximum shown people in search
  @property({
    attribute: 'show-max',
    type: Number
  })
  showMax: number = 6;

  //group attribute selection
  @property({
    attribute: 'group',
    type: String
  })
  group: string;

  //User selected people
  @property() public _selectedPeople: Array<any> = [];
  //single matching id for filtering against people list
  @property() private _duplicatePersonId: string = '';
  //User input in search
  @property() private _userInput: string = '';

  //tracking of user arrow key input for selection
  private arrowSelectionCount: number = 0;
  //List of people requested if group property is provided
  private groupPeople: any[];
  //if search is still loading don't load "people not found" state
  private isLoading = false;

  attributeChangedCallback(att, oldval, newval) {
    super.attributeChangedCallback(att, oldval, newval);

    if (att == 'group' && oldval !== newval) {
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
    if (this.group) {
      Providers.onProviderUpdated(() => this.findGroup());
      this.findGroup();
    }
  }

  private async findGroup() {
    let provider = Providers.globalProvider;
    if (provider && provider.state === ProviderState.SignedIn) {
      let client = Providers.globalProvider.graph;
      this.groupPeople = await client.getPeopleFromGroup(this.group);
    }
  }

  private onUserTypeSearch(event: any) {
    if (event.keyCode == 40 || event.keyCode == 38) {
      //keyCodes capture: down arrow (40) and up arrow (38)
      //logic is in onUserKeyDown
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
      this.fireCustomEvent('selectionChanged', this._selectedPeople);
      return;
    }
    this._userInput = event.target.value;
    if (event.target.value) {
      this.loadPersonSearch(this._userInput);
      window.addEventListener(
        'keyup',
        debounce(() => {
          this.loadPersonSearch(this._userInput);
        }, 300)
      );
    } else {
      event.target.value = '';
      this._userInput = '';
      this.people = [];
    }
  }

  private onUserKeyDown(event: any) {
    if (event.keyCode == 40 || event.keyCode == 38) {
      //keyCodes capture: down arrow (40) and up arrow (38)
      this.handleArrowSelection(event);
      event.preventDefault();
    }
    if (event.code == 'Tab' || event.code == 'Enter') {
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
        if (this.arrowSelectionCount + 1 !== this.people.length && this.arrowSelectionCount + 1 < this.showMax) {
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

  private async loadPersonSearch(name: string) {
    this.isLoading = true;
    if (name.length) {
      name = name.toLowerCase();
      let provider = Providers.globalProvider;
      let peoples: any;
      if (provider && provider.state === ProviderState.SignedIn) {
        let client = Providers.globalProvider.graph;
        //filtering groups
        if (this.group) {
          peoples = this.groupPeople.filter(function(person) {
            return person.displayName.toLowerCase().indexOf(name) !== -1;
          });
          for (let person of peoples) {
            // set image to @ to flag the mgt-person component to
            // query the image from the graph
            person.image = '@';
          }
          this.filterPeople(peoples);
          return;
        }
        peoples = await client.findPerson(name);
        for (let person of peoples) {
          // set image to @ to flag the mgt-person component to
          // query the image from the graph
          person.image = '@';
        }
        if (peoples) {
          peoples = peoples.filter(function(person) {
            return person.displayName.toLowerCase().indexOf(name) !== -1;
          });
          this.filterPeople(peoples);
        } else {
          this.people = [];
          this.filterPeople(peoples);
        }
      }
    }
  }

  private filterPeople(peoples: any) {
    //check if people need to be updated
    //ensuring people list is displayed
    //find ids from selected people
    if (peoples) {
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
    }
    this.isLoading = false;
  }

  private removePerson(person: MgtPersonDetails) {
    let chosenPerson: any = person;
    let filteredPersonArr = this._selectedPeople.filter(function(person) {
      return person.id !== chosenPerson.id;
    });
    this._selectedPeople = filteredPersonArr;
    this.fireCustomEvent('selectionChanged', this._selectedPeople);
    this.renderChosenPeople();
  }

  private renderErrorMessage() {
    return html`
      <div class="error-message-parent">
        <div class="search-error-text">We didn't find any matches.</div>
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
            .value="${this._userInput}"
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
      if (peoples.length == 0 && this._userInput.length > 0 && this.isLoading!) {
        return html`
          <span class="error-message-holder">${this.renderErrorMessage()}</span>
        `;
      } else {
        return html`
          <ul class="people-list">
            ${this.renderPersons(peoples)}
          </ul>
        `;
      }
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
