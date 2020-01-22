/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { templateElement } from '@babel/types';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property, TemplateResult } from 'lit-element';
import { isTemplatePartActive } from 'lit-html';
import { repeat } from 'lit-html/directives/repeat';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import '../../styles/fabric-icon-font';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';
import { debounce } from '../../utils/Utils';
import '../mgt-person/mgt-person';
import { MgtTemplatedComponent } from '../templatedComponent';
import { styles } from './mgt-teams-channel-picker-css';

/**
 * Web component used to select channels from a User's Microsoft Teams profile
 *
 *
 * @class MgtTeamsChannelPicker
 * @extends {MgtTemplatedComponent}
 */
@customElement('mgt-teams-channel-picker')
export class MgtTeamsChannelPicker extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * user the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  @property({
    attribute: 'teams',
    type: Object
  })
  public teams: any = null;

  @property({
    attribute: 'selected-teams'
  })
  public selectedTeams: any[] = [];

  // User input in search
  @property() private _userInput: string = '';

  @property() private _showChannel: any[] = [];

  @property() private _teamChannels: any[];

  private hideChannels: boolean = true;

  // tracking of user arrow key input for selection
  private arrowSelectionCount: number = 0;
  // if search is still loading don't load "people not found" state
  private isLoading = false;
  private debouncedSearch;

  constructor() {
    super();
  }

  /**
   * Invoked when the element is first updated. Implement to perform one time work on the element after update.
   *
   * @memberof MgtTeamsChannelPicker
   */
  public firstUpdated() {
    Providers.onProviderUpdated(() => {
      this.loadTeams();
    });
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return a lit-html TemplateResult.
   * Setting properties inside this method will not trigger the element to update.
   * @returns
   * @memberof MgtTeamsChannelPicker
   */
  public render() {
    return (
      this.renderTemplate('default', { channels: this._teamChannels }) ||
      html`
        <div class="people-picker" @blur=${this.lostFocus}>
          <div class="people-picker-input" @click=${this.gainedFocus}>
            ${this.renderChosenPeople()}
          </div>
          <div class="people-list-separator"></div>
          ${this.renderChannelList()}
        </div>
      `
    );
  }

  /**
   * Adds debounce method for set delay on user input
   */
  private onUserKeyUp(event: any) {
    if (event.keyCode === 40 || event.keyCode === 38) {
      // keyCodes capture: down arrow (40) and up arrow (38)
      return;
    }

    const input = event.target;

    if (event.code === 'Escape') {
      input.value = '';
      this._userInput = '';
      return;
    }
    if (event.code === 'Backspace' && this._userInput.length === 0 && this.selectedTeams.length > 0) {
      input.value = '';
      this._userInput = '';
      // remove last channel in selected list
      this.selectedTeams = this.selectedTeams.splice(0, this.selectedTeams.length - 1);
      // fire selected teams changed event
      this.fireCustomEvent('selectionChanged', this.selectedTeams);
      return;
    }

    this.handleChannelSearch(input);
  }

  /**
   * Tracks event on user input in search
   * @param input - input text
   */
  private handleChannelSearch(input: any) {
    if (!this.debouncedSearch) {
      this.debouncedSearch = debounce(() => {
        if (this._userInput !== input.value) {
          this._userInput = input.value;
          this.loadChannelSearch(this._userInput);
          this.arrowSelectionCount = 0;
        }
      }, 200);
    }

    this.debouncedSearch();
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
      if (this._teamChannels.length) {
        event.preventDefault();
      }

      event.target.value = '';
    }
  }

  /**
   * Tracks user key selection for arrow key selection of channels
   * @param event - tracks user key selection
   */
  private handleArrowSelection(event: any) {
    if (this._teamChannels.length) {
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
        if (this.arrowSelectionCount + 1 !== this._teamChannels.length && this.arrowSelectionCount + 1) {
          this.arrowSelectionCount++;
        } else {
          this.arrowSelectionCount = 0;
        }
      }

      const channelList = this.renderRoot.querySelector('.people-list');
      // reset background color
      // tslint:disable-next-line: prefer-for-of
      for (let i = 0; i < channelList.children.length; i++) {
        channelList.children[i].setAttribute('class', 'list-person people-person-list');
      }
      // set selected background
      channelList.children[this.arrowSelectionCount].setAttribute('class', 'list-person people-person-list-fill');
    }
  }

  private async loadTeams() {
    const provider = Providers.globalProvider;
    const client = Providers.globalProvider.graph;

    this.teams = await client.picker_GetAllMyTeams();

    console.log('teams', this.teams);

    if (provider) {
      if (provider.state === ProviderState.SignedIn) {
        const batch = provider.graph.createBatch();

        for (const [i, team] of this.teams.entries()) {
          batch.get(`${i}`, `teams/${team.id}/channels`, ['user.read.all']);
        }
        const response = await batch.execute();
        this._teamChannels = this.teams.map((team, index) => [team, response[index].value]);
      }
    }
    console.log('teamChannels', this._teamChannels);
  }

  /**
   * Async method which query's the Graph with user input
   * @param name - user input or name of person searched
   */
  private loadChannelSearch(name: string) {
    if (name !== '') {
      this.hideChannels = false;
    } else {
      this.hideChannels = true;
    }
    for (let i = 0; i < this._teamChannels.length; i++) {
      for (let j = 0; j < this._teamChannels[i][1].length; j++) {
        if (this._teamChannels[i][1][j] && this._teamChannels[i][1][j].displayName.toLowerCase().indexOf(name) !== -1) {
          this._teamChannels[i][1][j].isShowing = true;
        } else if (
          this._teamChannels[i][1][j] &&
          this._teamChannels[i][1][j].displayName.toLowerCase().indexOf(name) === -1
        ) {
          this._teamChannels[i][1][j].isShowing = false;
        }
      }
    }
  }

  /**
   * Removes person from selected people
   * @param person - person and details pertaining to user selected
   */
  private removePerson(person: MicrosoftGraph.User | MicrosoftGraph.Person | MicrosoftGraph.Contact) {
    const chosenPerson: any = person;
    const filteredPersonArr = this.selectedTeams.filter(p => {
      return p.id !== chosenPerson.id;
    });
    this.selectedTeams = filteredPersonArr;
    this.fireCustomEvent('selectionChanged', this.selectedTeams);
  }

  private renderErrorMessage() {
    return html`
      <div class="message-parent">
        <div label="search-error-text" aria-label="We didn't find any matches." class="search-error-text">
          We didn't find any matches.
        </div>
      </div>
    `;
  }

  private renderLoadingMessage() {
    return html`
      <div class="message-parent">
        <div label="search-error-text" aria-label="loading" class="loading-text">
          ......
        </div>
      </div>
    `;
  }

  private renderChosenPeople() {
    let peopleList;
    let inputClass = 'input-search-start';

    if (this.selectedTeams.length > 0) {
      inputClass = 'input-search';

      peopleList = html`
        ${this.selectedTeams.slice(0, this.selectedTeams.length).map(
          team =>
            html`
              <li class="people-person">
                ${team[0].displayName + ':' + team[1].displayName}
                <div class="CloseIcon" @click="${() => this.removePerson(team)}"></div>
              </li>
            `
        )}
      `;
    } else {
      peopleList = null;
    }
    // tslint:disable
    return html`
      <div class="people-chosen-list">
        ${peopleList}
        <div class="${inputClass}">
          <input
            id="people-picker-input"
            class="people-chosen-input ${this.selectedTeams.length > 0 ? 'hide' : ''}"
            type="text"
            placeholder="Select a channel: "
            label="people-picker-input"
            aria-label="people-picker-input"
            role="input"
            .value="${this._userInput}"
            @keydown="${this.onUserKeyDown}"
            @keyup="${this.onUserKeyUp}"
          />
        </div>
      </div>
    `;
  }
  // tslint:enable

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
      peopleList.setAttribute('style', 'display:none');
    }
  }

  private renderHighlightText(channel: any) {
    const channels: any = channel;

    const highlightLocation = channels.displayName.toLowerCase().indexOf(this._userInput.toLowerCase());
    if (highlightLocation !== -1) {
      // no location
      if (highlightLocation === 0) {
        // highlight is at the beginning of sentence
        channels.first = '';
        channels.highlight = channels.displayName.slice(0, this._userInput.length);
        channels.last = channels.displayName.slice(this._userInput.length, channels.displayName.length);
      } else if (highlightLocation === channels.displayName.length) {
        // highlight is at end of the sentence
        channels.first = channels.displayName.slice(0, highlightLocation);
        channels.highlight = channels.displayName.slice(highlightLocation, channels.displayName.length);
        channels.last = '';
      } else {
        // highlight is in middle of sentence
        channels.first = channels.displayName.slice(0, highlightLocation);
        channels.highlight = channels.displayName.slice(highlightLocation, highlightLocation + this._userInput.length);
        channels.last = channels.displayName.slice(
          highlightLocation + this._userInput.length,
          channels.displayName.length
        );
      }
    }

    return html`
      <div>
        <span class="people-person-text">${channels.first}</span
        ><span class="people-person-text highlight-search-text">${channels.highlight}</span
        ><span class="people-person-text">${channels.last}</span>
      </div>
    `;
  }

  private renderChannelList() {
    let content: TemplateResult;

    if (this._teamChannels) {
      content = this.renderTeams(this._teamChannels);

      if (this.isLoading) {
        content = this.renderTemplate('loading', null, 'loading') || this.renderLoadingMessage();
      } else if (this._teamChannels.length === 0 && this._userInput.length > 0) {
        content = this.renderTemplate('error', null, 'error') || this.renderErrorMessage();
      } else {
        if (this._teamChannels[0]) {
          (this._teamChannels[0] as any).isSelected = 'fill';
        }
      }
    }

    return html`
      <div class="people-list">
        ${content}
      </div>
    `;
  }

  private async loadChannels(team: any) {
    this._showChannel.push(team.id);
  }
  private renderChannels(team: any[], channels: any[]) {
    let channelView;
    channelView = html`
      ${repeat(
        channels,
        channel => channel,
        channel => html`
          <div class="channel-display" @click=${() => this.addChannel(team, channel)}>
            ${channel.isShowing !== false ? this.renderHighlightText(channel) : null}
          </div>
        `
      )}
    `;

    return channelView;
  }

  private addChannel(team, channel) {
    this.selectedTeams.push([team, channel]);
    this.lostFocus();

    this.requestUpdate();
  }

  private _clickTeam(channel: any) {
    const team = this.renderRoot.querySelector('.team-' + channel.id);

    if (team.classList[2] === 'hide-channels') {
      team.classList.remove('hide-channels');
    } else {
      team.classList.add('hide-channels');
    }
  }

  private renderTeam(team) {
    window.setTimeout(() => {
      const channelsDiv = this.renderRoot.querySelector('.team-' + team.id);
      let noChannelsDisplayed = true;
      if (channelsDiv) {
        for (let i = 0; i < channelsDiv.children.length; i++) {
          if (channelsDiv.children[i].children.length > 0) {
            noChannelsDisplayed = false;
          }
        }
      }

      const teamDiv = this.renderRoot.querySelector('.teamName-' + team.displayName.replace(/\s/g, ''));
      if (teamDiv) {
        teamDiv.parentElement.classList.remove('hide-team');
        if (noChannelsDisplayed && this._userInput.length) {
          teamDiv.parentElement.classList.add('hide-team');
        } else {
          teamDiv.parentElement.classList.remove('hide-team');
        }
      }
    });
  }

  private renderTeams(teamChannels: any[]) {
    return html`
      ${repeat(
        teamChannels,
        teamData => teamData[0].id,
        teamData => html`
          <li class="list-person people-person-list" @click="${() => this.loadChannels(teamData)}">
            ${this._showChannel[0] === teamData[0].id
              ? getSvg(SvgIcon.ArrowDown, '#252424')
              : getSvg(SvgIcon.ArrowRight, '#252424')}
            ${this.renderTeam(teamData[0])}
            <div
              class=${'people-person-text-area teamName-' + teamData[0].displayName.replace(/\s/g, '')}
              id="${teamData[0].displayName.replace(/\s/g, '')}"
              @click=${() => this._clickTeam(teamData[0])}
            >
              ${teamData[0].displayName}
            </div>
          </li>
          <div class="render-channels team-${teamData[0].id} ${this.hideChannels ? 'hide-channels' : ''}">
            ${this.renderChannels(teamData[0], teamData[1])}
          </div>
        `
      )}
    `;
  }
}
