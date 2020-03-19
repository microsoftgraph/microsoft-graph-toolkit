/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { forStatement, templateElement } from '@babel/types';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property, TemplateResult } from 'lit-element';
import { isTemplatePartActive } from 'lit-html';
import { classMap } from 'lit-html/directives/class-map';
import { repeat } from 'lit-html/directives/repeat';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import '../../styles/fabric-icon-font';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';
import { debounce } from '../../utils/Utils';
import '../mgt-person/mgt-person';
import { MgtTemplatedComponent } from '../templatedComponent';
import { styles } from './mgt-teams-channel-picker-css';
import { getAllMyTeams } from './mgt-teams-channel-picker.graph';

/**
 * Drop down menu item
 *
 * @export
 * @interface DropdownItem
 */
export interface DropdownItem {
  /**
   * Teams channel
   *
   * @type {DropdownItem[]}
   * @memberof DropdownItem
   */
  channels?: DropdownItem[];
  /**
   * Name of team or channel
   *
   * @type {string}
   * @memberof DropdownItem
   */
  displayName?: string;
}

/**
 * Drop down menu item state
 *
 * @interface DropdownItemState
 */
interface DropdownItemState {
  /**
   * provided dropdown item
   *
   * @type {DropdownItem}
   * @memberof DropdownItemState
   */
  item: DropdownItem;
  /**
   * if dropdown item shows expanded state
   *
   * @type {boolean}
   * @memberof DropdownItemState
   */
  isExpanded?: boolean;
  /**
   * If item contains channels
   *
   * @type {DropdownItemState[]}
   * @memberof DropdownItemState
   */
  channels?: DropdownItemState[];
  /**
   * if Item has parent item (team)
   *
   * @type {DropdownItemState}
   * @memberof DropdownItemState
   */
  parent: DropdownItemState;
}
/**
 * Establishes Microsoft Teams channels for use in Microsoft.Graph.Team type
 * @type MicrosoftGraph.Team
 *
 */

type Team = MicrosoftGraph.Team & {
  /**
   * Display name Of Team
   *
   * @type {string}
   */
  displayName: string;
  /**
   * Microsoft Graph Channel
   *
   * @type {MicrosoftGraph.Channel[]}
   */
  channels: MicrosoftGraph.Channel[];
  /**
   * Determines whether Team displays channels
   *
   * @type {Boolean}
   */
  showChannels: boolean;
};

/**
 * Establishes connection of MicrosoftGraph.Channel to its respecive Team
 * @type MicrosoftGraph.Channel
 *
 */

type Channel = MicrosoftGraph.Channel & {
  /**
   * Determines whether Team displays channels
   *
   * @type MicrosoftGraph.Team
   */
  Team: MicrosoftGraph.Team;
};

/**
 * Web component used to select channels from a User's Microsoft Teams profile
 *
 *
 * @class MgtTeamsChannelPicker
 * @extends {MgtTemplatedComponent}
 *
 * @cssprop --font-color - {font} Default font color
 *
 * @cssprop --input-border - {String} Input section entire border
 * @cssprop --input-border-top - {String} Input section border top only
 * @cssprop --input-border-right - {String} Input section border right only
 * @cssprop --input-border-bottom - {String} Input section border bottom only
 * @cssprop --input-border-left - {String} Input section border left only
 * @cssprop --input-background-color - {Color} Input section background color
 * @cssprop --input-hover-color - {Color} Input text hover color
 *
 * @cssprop --selection-background-color - {Color} Highlight of selected channel color
 * @cssprop --selection-hover-color - {Color} Highlight of selected channel color during hover state
 *
 * @cssprop --arrow-fill - {Color} Color of arrow svg
 * @cssprop --placeholder-focus-color - {Color} Highlight of placeholder text during focus state
 *
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

  // public property for setting the items of the dropdown

  /**
   * Sets Team items to be used
   *
   * @memberof MgtTeamsChannelPicker
   */
  public set items(value) {
    this._items = value;
    this._treeViewState = this.generateTreeViewState(value);
    this.resetFocusState();
  }
  public get items(): DropdownItem[] {
    return this._items;
  }

  /**
   * Gets Selected item to be used
   *
   * @readonly
   * @type {DropdownItem}
   * @memberof MgtTeamsChannelPicker
   */
  public get selectedItem(): DropdownItem {
    return this._selectedItemState ? this._selectedItemState.item : null;
  }

  /**
   * user's Microsoft joinedTeams
   *
   * @type {Team[]}
   * @memberof MgtTeamsChannelPicker
   */
  @property({
    attribute: 'teams',
    type: Object
  })
  public teams: Team[] = [];

  /**
   * user selected teams
   *
   * @type {any[]}
   * @memberof MgtTeamsChannelPicker
   */
  @property({
    attribute: 'selected-teams'
  })
  public selectedTeams: [Team[], MicrosoftGraph.Channel[]] = [[], []];

  /**
   *  array of user picked people.
   * @type {Array<any>}
   */
  @property({
    attribute: 'selected-channel',
    type: Array
  })
  public selectedChannel: Channel[] = [];

  // User input in search
  private _userInput: string = '';

  // tracking of user arrow key input for selection
  private arrowSelectionCount: number = -1;

  private channelLength: number = 0;

  private channelCounter: number = -1;

  private isHovered = false;

  private isFocused = false;

  // Properties

  private _selectedItemState: DropdownItemState;
  private _items: DropdownItem[] = [];
  private _treeViewState: DropdownItemState[] = [];

  // focus state
  private _focusList = [];
  private _focusedIndex: number = -1;

  // determines loading state
  @property() private isLoading = false;
  private debouncedSearch;

  constructor() {
    super();
    this.handleWindowClick = this.handleWindowClick.bind(this);
  }

  /**
   * Invoked each time the custom element is appended into a document-connected element
   *
   * @memberof MgtPerson
   */
  public connectedCallback() {
    super.connectedCallback();
    window.addEventListener('click', this.handleWindowClick);
  }

  /**
   * Invoked each time the custom element is disconnected from the document's DOM
   *
   * @memberof MgtPerson
   */
  public disconnectedCallback() {
    window.removeEventListener('click', this.handleWindowClick);
    super.disconnectedCallback();
  }

  /**
   * Queries the microsoft graph for a user based on the user id and adds them to the selectedPeople array
   *
   * @param {[string]} an array of user ids to add to selectedPeople
   * @returns {Promise<void>}
   * @memberof MgtPeoplePicker
   */
  public async selectChannelsById(channelIds: [string]): Promise<void> {
    const provider = Providers.globalProvider;
    const client = Providers.globalProvider.graph;
    if (provider && provider.state === ProviderState.SignedIn) {
      for (const team of this.teams) {
        for (const channel of team.channels) {
          if (channel.id === channelIds[0]) {
            try {
              this.selectedTeams = [[team], [channel]];
              // tslint:disable-next-line: no-empty
            } catch (e) {}
          }
        }
      }
    }
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return a lit-html TemplateResult.
   * Setting properties inside this method will not trigger the element to update.
   * @returns
   * @memberof MgtTeamsChannelPicker
   */
  public render() {
    return (
      this.renderTemplate('default', { teams: this.teams }) ||
      html`
        <div
          class="teams-channel-picker"
          @mouseover=${this.handleHover}
          @blur=${this.lostFocus}
          @mouseout=${this.mouseLeft}
        >
          <div
            class="teams-channel-picker-input ${this.isHovered ? 'hovered' : ''} ${this.isFocused ? 'focused' : ''}"
            @click=${this.gainedFocus}
          >
            <div class="search-icon ${this.isFocused && this.selectedTeams[0].length === 0 ? 'focused' : ''}">
              ${getSvg(SvgIcon.Search, '#252424')}
            </div>
            ${this.renderChosenTeam()}
          </div>
          <div class="team-list">${this.renderDropdown(this._treeViewState)}</div>
        </div>
      `
    );
  }

  /**
   * Renders list of teams
   *
   * @param {DropdownItemState[]} items
   * @param {number} [level=0]
   * @returns
   * @memberof MgtTeamsChannelPicker
   */
  public renderDropdown(items: DropdownItemState[], level: number = 0) {
    let content: any;
    if (this.teams) {
      if (this.isLoading) {
        content = this.renderTemplate('loading', null, 'loading') || this.renderLoadingMessage();
      } else {
        content = items.map((treeItem, index) => {
          const isLeaf = !treeItem.channels;
          const renderChannels = !isLeaf && treeItem.isExpanded;

          return html`
            ${this.renderItem(treeItem)} ${renderChannels ? this.renderDropdown(treeItem.channels, level + 1) : html``}
          `;
        });
        this.requestUpdate();
      }
    }

    return html`
      <div>
        ${content}
      </div>
    `;
  }

  /**
   * Renders each Channel or Team
   *
   * @param {DropdownItemState} itemState
   * @returns
   * @memberof MgtTeamsChannelPicker
   */
  public renderItem(itemState: DropdownItemState) {
    let icon: TemplateResult = null;

    if (itemState.channels) {
      // must be team with channels
      icon = itemState.isExpanded ? getSvg(SvgIcon.ArrowDown, '#252424') : getSvg(SvgIcon.ArrowRight, '#252424');
    }

    const classes = {
      focused: this._focusList[this._focusedIndex] === itemState,
      item: true,
      listTeam: itemState.channels ? true : false,
      selected: this.selectedItem === itemState.item
    };

    return html`
      <div @click=${() => this.handleItemClick(itemState)} class="${classMap(classes)}">
        <div class="arrow">
          ${icon}
        </div>
        ${itemState.channels ? itemState.item.displayName : this.renderHighlightText(itemState.item)}
      </div>
    `;
  }

  /**
   *  handles clicking of dropdown menu item and resets list state
   *
   * @param {DropdownItemState} item
   * @memberof MgtTeamsChannelPicker
   */
  public handleItemClick(item: DropdownItemState) {
    if (item.channels) {
      item.isExpanded = !item.isExpanded;
    } else if (this._selectedItemState !== item) {
      this._selectedItemState = item;
    } else {
      this._selectedItemState = null;
    }

    this._focusedIndex = -1;
    this.resetFocusState();
  }

  /**
   * Updates list state with items if changed
   *
   * @param {*} e
   * @memberof MgtTeamsChannelPicker
   */
  public handleInputChanged(e) {
    if (this._userInput !== e.target.value) {
      this._userInput = e.target.value;
    }
    if (!this.debouncedSearch) {
      this.debouncedSearch = debounce(() => {
        this._treeViewState = this.generateTreeViewState(this.items, this._userInput);
        this._focusedIndex = -1;
        this.resetFocusState();
      }, 800);
    }

    this.debouncedSearch();
  }

  /**
   * Renders selected Team
   *
   * @protected
   * @returns
   * @memberof MgtTeamsChannelPicker
   */
  protected renderChosenTeam() {
    let channelList;
    let inputClass = 'input-search-start';

    if (this.selectedTeams[0]) {
      if (this.selectedTeams[0].length) {
        inputClass = 'input-search';

        channelList = html`
          <li class="selected-team">
            <div class="selected-team-name">${this.selectedTeams[0][0].displayName}</div>
            <div class="arrow">${getSvg(SvgIcon.TeamSeparator, '#B3B0AD')}</div>
            ${this.selectedTeams[1][0].displayName}
            <div class="CloseIcon" @click="${() => this.removeTeam(this.selectedTeams[0], this.selectedTeams[1])}">
              
            </div>
          </li>
          <div class="SearchIcon">
            ${getSvg(SvgIcon.Search, '#252424')}
          </div>
        `;
      } else {
        channelList = html`
          <div
            class="InputArrowIcon"
            @click="${e => {
              e.stopPropagation();
              this.isFocused ? this.lostFocus() : this.gainedFocus();
            }}"
          >
            ${this.isFocused ? getSvg(SvgIcon.UpCarrot, '#605E5C') : getSvg(SvgIcon.DownCarrot, '#605E5')}
          </div>
        `;
      }
    }
    return html`
      <div class="people-chosen-list">
        ${channelList}
        <div class="${inputClass}">
          <input
            id="teams-channel-picker-input"
            class="team-chosen-input ${this.isFocused || this.isHovered ? 'focused' : ''}"
            type="text"
            placeholder="${this.selectedTeams[0].length > 0 ? '' : 'Select a channel '} "
            label="teams-channel-picker-input"
            aria-label="teams-channel-picker-input"
            role="input"
            .value="${this._userInput}"
            @keydown="${this.onUserKeyDown}"
            @input=${e => this.handleInputChanged(e)}
          />
        </div>
      </div>
    `;
  }

  private generateTreeViewState(
    tree: DropdownItem[],
    filterString: string = '',
    parent: DropdownItemState = null
  ): DropdownItemState[] {
    const treeView: DropdownItemState[] = [];
    filterString = filterString.toLowerCase();

    for (const item of tree) {
      let stateItem: DropdownItemState;

      if (filterString.length === 0 || item.displayName.toLowerCase().includes(filterString)) {
        stateItem = { item, parent };
        if (item.channels) {
          stateItem.channels = this.generateTreeViewState(item.channels, '', stateItem);
          stateItem.isExpanded = filterString.length > 0;
        }
      } else if (item.channels) {
        const channels = this.generateTreeViewState(item.channels, filterString, stateItem);
        if (channels.length > 0) {
          stateItem = { item, parent, channels, isExpanded: true };
        }
      }

      if (stateItem) {
        treeView.push(stateItem);
      }
    }

    return treeView;
  }

  // generates a flat list from a tree to facilitate easier focus
  // navigation
  private generateFocusList(items): any[] {
    if (!items || items.length === 0) {
      return [];
    }

    let array = [];

    for (const item of items) {
      array.push(item);
      if (item.channels && item.isExpanded) {
        array = [...array, ...this.generateFocusList(item.channels)];
      }
    }

    return array;
  }

  private resetFocusState() {
    this._focusList = this.generateFocusList(this._treeViewState);
    this.requestUpdate();
  }

  private mouseLeft() {
    if (this._userInput.length === 0) {
      this.isHovered = false;
      this.requestUpdate();
    }
  }

  private handleHover() {
    if (this.teams.length === 0) {
      this.loadTeams();
    }
    this.isHovered = true;
    this.requestUpdate();
  }

  private handleWindowClick(e: MouseEvent) {
    if (e.target !== this) {
      this.lostFocus();
    }
  }

  /**
   * Tracks event on user search (keydown)
   * @param event - event tracked on user input (keydown)
   */
  private onUserKeyDown(event: any) {
    if (this._treeViewState.length === 0) {
      return;
    }

    const currentFocusedItem = this._focusList[this._focusedIndex];

    switch (event.keyCode) {
      case 40: // up
        this._focusedIndex = (this._focusedIndex + 1) % this._focusList.length;
        if (typeof this._focusedIndex !== 'number') {
          this._focusedIndex = -1;
        }

        this.requestUpdate();
        event.preventDefault();
        break;
      case 38: // down
        this._focusedIndex = (this._focusedIndex - 1 + this._focusList.length) % this._focusList.length;
        if (typeof this._focusedIndex !== 'number') {
          this._focusedIndex = -1;
        }

        this.requestUpdate();
        event.preventDefault();
        break;
      case 39: // right
        if (currentFocusedItem && currentFocusedItem.channels && !currentFocusedItem.isExpanded) {
          currentFocusedItem.isExpanded = true;
          this.resetFocusState();
          event.preventDefault();
        }
        break;
      case 37: // left
        if (currentFocusedItem && currentFocusedItem.channels && currentFocusedItem.isExpanded) {
          currentFocusedItem.isExpanded = false;
          this.resetFocusState();
          event.preventDefault();
        }
        break;
    }
  }

  private async loadTeams() {
    const provider = Providers.globalProvider;
    if (provider) {
      if (provider.state === ProviderState.SignedIn) {
        const graph = provider.graph.forComponent(this);

        this.teams = await getAllMyTeams(graph);

        this.isLoading = true;
        const batch = provider.graph.createBatch();

        for (const [i, team] of this.teams.entries()) {
          batch.get(`${i}`, `teams/${team.id}/channels`, ['user.read.all']);
        }
        const response = await batch.execute();
        for (const [i, team] of this.teams.entries()) {
          team.channels = response[i].value;
        }
      }
    }
    this.items = this.teams;
    this.resetFocusState();
    this.isLoading = false;
  }

  /**
   * Async method which query's the Graph with user input
   * @param name - user input or name of person searched
   */
  private loadChannelSearch(name: string) {
    if (name === '') {
      return;
    }
    const foundMatch = [];
    for (const team of this.teams) {
      for (const channel of team.channels) {
        if (channel.displayName.toLowerCase().indexOf(name) !== -1) {
          foundMatch.push(team.displayName);
        }
      }
    }

    if (foundMatch.length) {
      for (const team of this.teams) {
        const teamdiv = this.renderRoot.querySelector(`.team-list-${team.id}`);
        if (teamdiv) {
          teamdiv.classList.remove('hide-team');
          team.showChannels = true;
          if (foundMatch.indexOf(team.displayName) === -1) {
            team.showChannels = false;
            teamdiv.classList.add('hide-team');
          }
        }
      }
    }
  }

  private removeTeam(team: MicrosoftGraph.Team[], pickedChannel: MicrosoftGraph.Channel[]) {
    this.selectedTeams = [[], []];
    this._userInput = '';

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
        <div class="spinner"></div>
        <div label="loading-text" aria-label="loading" class="loading-text">
          Loading...
        </div>
      </div>
    `;
  }

  private gainedFocus() {
    this.isFocused = true;
    const teamList = this.renderRoot.querySelector('.team-list');
    const teamInput = this.renderRoot.querySelector('.team-chosen-input') as HTMLInputElement;
    teamInput.focus();

    if (teamList) {
      // Mouse is focused on input
      teamList.setAttribute('style', 'display:block');
    }
    this.requestUpdate();
  }

  private lostFocus() {
    this.isFocused = false;
    this._userInput = '';
    const teamList = this.renderRoot.querySelector('.team-list');
    if (teamList) {
      teamList.setAttribute('style', 'display:none');
    }
    this.requestUpdate();
  }

  private renderHighlightText(channel: any) {
    // tslint:disable-next-line: prefer-const
    let channels: any = {};

    let shouldShow = true;

    const highlightLocation = channel.displayName.toLowerCase().indexOf(this._userInput.toLowerCase());
    if (highlightLocation !== -1) {
      // no location
      if (highlightLocation === 0) {
        // highlight is at the beginning of sentence
        channels.first = '';
        channels.highlight = channel.displayName.slice(0, this._userInput.length);
        channels.last = channel.displayName.slice(this._userInput.length, channel.displayName.length);
      } else if (highlightLocation === channel.displayName.length) {
        // highlight is at end of the sentence
        channels.first = channel.displayName.slice(0, highlightLocation);
        channels.highlight = channel.displayName.slice(highlightLocation, channel.displayName.length);
        channels.last = '';
      } else {
        // highlight is in middle of sentence
        channels.first = channel.displayName.slice(0, highlightLocation);
        channels.highlight = channel.displayName.slice(highlightLocation, highlightLocation + this._userInput.length);
        channels.last = channel.displayName.slice(
          highlightLocation + this._userInput.length,
          channel.displayName.length
        );
      }
    } else {
      shouldShow = false;
    }

    return html`
      <div class="showing channel-display">
        <span class="people-person-text">${channels.first}</span
        ><span class="people-person-text highlight-search-text">${channels.highlight}</span
        ><span class="people-person-text">${channels.last}</span>
      </div>
    `;
  }

  private addChannel(event, pickedChannel: any) {
    // reset blue highlight

    for (const team of this.teams) {
      // tslint:disable-next-line: prefer-for-of
      for (let i = 0; i < team.channels.length; i++) {
        const selection = team.channels[i].id.replace(/[^a-zA-Z ]/g, '');
        const channelDiv = this.renderRoot.querySelector(`.channel-${selection}`);
        channelDiv.parentElement.classList.remove('blue-highlight');
      }
    }

    if (event.key === 'Tab') {
      for (const team of this.teams) {
        if (team.id === pickedChannel) {
          const selection = team.channels[this.channelCounter].id.replace(/[^a-zA-Z ]/g, '');
          const channelDiv = this.renderRoot.querySelector(`.channel-${selection}`);

          const shownIds = [];

          // check if channels are filtered
          // tslint:disable-next-line: prefer-for-of
          for (let i = 0; i < channelDiv.parentElement.parentElement.children.length; i++) {
            if (channelDiv.parentElement.parentElement.children[i].children[0].classList.contains('showing')) {
              shownIds.push(channelDiv.parentElement.parentElement.children[i].children[0].classList[1].slice(8));
            }
          }

          for (const channel of team.channels) {
            if (channel.id.replace(/[^a-zA-Z ]/g, '') === shownIds[this.channelCounter]) {
              this.selectedTeams = [[team], [channel]];
            }
          }

          channelDiv.parentElement.classList.add('blue-highlight');
          this.arrowSelectionCount = -1;
          this.lostFocus();
        }
      }
    } else {
      const teamDiv =
        event.target.parentNode.parentNode.classList[1] || event.target.parentNode.parentNode.parentNode.classList[1];
      if (teamDiv) {
        const teamId = teamDiv.slice(5, teamDiv.length);
        for (const team of this.teams) {
          if (team.id === teamId) {
            this.selectedTeams = [[team], [pickedChannel]];
            const selection = pickedChannel.id.replace(/[^a-zA-Z ]/g, '');
            const channelDiv = this.renderRoot.querySelector(`.channel-${selection}`);
            channelDiv.parentElement.classList.add('blue-highlight');
            this.lostFocus();
          }
        }
      }
    }
    this._userInput = '';
    this.arrowSelectionCount = -1;
    this.channelCounter = 0;
    this.requestUpdate();
  }
}
