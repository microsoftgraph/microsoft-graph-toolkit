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

  @property() private _personName: string = '';
  @property() private _selectedPeople: Array<any> = [];
  @property() private _duplicatePersonId: string = '';
  @property() private _userInput: string = '';
  @property() private _previousSearch: any;

  @property() private arrowSelectionCount: number = 0;
  /* TODO: Do we want a query property for loading groups from calls? */

  static get styles() {
    return styles;
  }

  constructor() {
    super();
  }

  private onUserTypeSearch(event: any) {
    if (event.code == 'Escape') {
      event.target.value = '';
      this._userInput = '';
      this.people = [];
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
    if (event.keyCode == 40 || event.keyCode == 38) {
      this.handleArrowSelection(event);
    }
  }

  private onUserKeyDown(event: any) {
    if (event.code == 'Tab') {
      this.addPerson(this.people[0], event);
    }
  }

  private handleArrowSelection(event: any) {
    let peoples: any = this.people;
    console.log('arrow selection count', this.arrowSelectionCount);
    for (let i = 0; i < peoples.length; i++) {
      peoples[i].isSelected = '';
    }
    if (peoples[this.arrowSelectionCount]) {
      peoples[this.arrowSelectionCount].isSelected = 'fill';
    } else {
      peoples[0].isSelected = 'fill';
      this.arrowSelectionCount = 0;
    }
    this.renderPersons(peoples);
    this.arrowSelectionCount++;
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
        let event = new CustomEvent('selectedPeople', {
          bubbles: true,
          cancelable: false,
          detail: this._selectedPeople
        });
        this.dispatchEvent(event);

        this.people = [];
        this._userInput = '';
        this._personName = '';
      }
    }
  }

  private async loadPersonSearch(name: string) {
    let provider = Providers.globalProvider;

    if (provider && provider.state === ProviderState.SignedIn) {
      let client = Providers.globalProvider.graph;

      let peoples: any = await client.findPerson(name);

      this.filterPeople(peoples);
    }
  }

  private filterPeople(peoples: any) {
    //check if people need to be updated
    if (this.people) {
      if (this.people.length > 0) this._previousSearch = this.people;
      //find ids from previous search
      let id_filter = peoples.map(function(el) {
        return el.id;
      });
      var filtered = this._previousSearch.filter(function(person) {
        return id_filter.indexOf(person.id) === -1;
      });
      if (filtered.length == 0 && this._userInput.length > 0) {
        return;
      } else {
        this.people = peoples;
      }
    } else {
      this.people = peoples;
    }
    //filter already selected people
    let selected = this._selectedPeople;
    for (let i = 0; i < selected.length; i++) {
      this.people = peoples.filter(function(person: any) {
        return person.id !== selected[i].id;
      });
    }
    for (var i = 0; i < this.people.length; i++) {
      if (peoples[i].image == undefined) {
        this.updateProfile(this.people);
      }
    }
  }

  private async updateProfile(peoples: any) {
    let provider = Providers.globalProvider;
    if (this.people) {
      for (var i = 0; i < peoples.length; i++) {
        if (peoples[i].id && peoples[i].image == undefined) {
          await Promise.all([
            provider.graph.getUser(peoples[i].id).then(user => {
              if (user) {
                peoples[i].displayName = user.displayName;
                this.requestUpdate();
              }
            }),
            provider.graph.getUserPhoto(peoples[i].id).then(photo => {
              if (photo) {
                peoples[i].image = photo;
                this.requestUpdate();
              }
            })
          ]);
        }
      }
    }
  }

  private removePerson(person: MgtPersonDetails) {
    let chosenPerson: any = person;
    let filteredPersonArr = this._selectedPeople.filter(function(person) {
      return person.id !== chosenPerson.id;
    });
    this._selectedPeople = filteredPersonArr;
    this.renderChosenPeople();
  }

  private renderErrorMessage() {
    if (this.people) {
      if (this.people.length == 0 && this._userInput.length > 0) {
        return html`
          <div class="error-message-parent">
            <div class="search-error-text">We didn't find any matches.</div>
          </div>
        `;
      }
    }
  }

  private renderChosenPeople() {
    if (this._selectedPeople.length > 0) {
      return html`
        <ul class="people-chosen-list">
          ${this._selectedPeople.slice(0, this._selectedPeople.length).map(
            person =>
              html`
                <li
                  class="${person.id == this._duplicatePersonId ? 'people-person duplicate-person' : 'people-person'}"
                >
                  ${this.renderTemplate('person', { person: person }, person.displayName) || this.renderPerson(person)}
                  <p class="person-display-name">${person.displayName}</p>
                  <div class="CloseIcon" @click="${() => this.removePerson(person)}">\uE711</div>
                </li>
              `
          )}
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
    } else {
      return html`
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
      `;
    }
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
      <p>
        <span class="people-person-text">${peoples.first}</span
        ><span class="people-person-text highlight-search-text">${peoples.highlight}</span
        ><span class="people-person-text">${peoples.last}</span>
      </p>
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
    console.log('render persns is called', peoples);
    return peoples.slice(0, this.showMax).map(
      person =>
        html`
          <li
            class="${peoples.isSelected == 'fill' ? 'people-person-list-fill' : 'people-person-list'}"
            @click="${(event: any) => this.addPerson(person, event)}"
          >
            ${this.renderTemplate('person', { person: person }, person.displayName) || this.renderPerson(person)}
            <p class="people-person-text-area" id="${person.displayName}">
              ${this.renderHighlightText(person)}
            </p>
          </li>
        `
    );
  }

  render() {
    return (
      this.renderTemplate('default', { people: this.people }) ||
      html`
        <div class="people-picker">
          <div class="people-picker-input">
            ${this.renderChosenPeople()}
          </div>
          <div class="people-list-separator"></div>
          ${this.renderPeopleList()}
          <div class="error-message-holder">
            ${this._userInput.length !== 0 ? this.renderErrorMessage() : null}
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
}
