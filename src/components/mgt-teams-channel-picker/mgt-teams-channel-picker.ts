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
import { MgtTemplatedComponent } from '../templatedComponent';
import { styles } from './mgt-teams-channel-picker-css';
import { getAllMyTeams } from './mgt-teams-channel-picker.graph';

/**
 * Team with displayName
 *
 * @export
 * @interface SelectedChannel
 */
export type Team = MicrosoftGraph.Team & {
  /**
   * Display name Of Team
   *
   * @type {string}
   */
  displayName?: string;
};

/**
 * Selected Channel item
 *
 * @export
 * @interface SelectedChannel
 */
export interface SelectedChannel {
  /**
   * Channel
   *
   * @type {MicrosoftGraph.Channel}
   * @memberof SelectedChannel
   */
  channel: MicrosoftGraph.Channel;

  /**
   * Team
   *
   * @type {MicrosoftGraph.Team}
   * @memberof SelectedChannel
   */
  team: Team;
}

/**
 * Drop down menu item
 *
 * @export
 * @interface DropdownItem
 */
interface DropdownItem {
  /**
   * Teams channel
   *
   * @type {DropdownItem[]}
   * @memberof DropdownItem
   */
  channels?: DropdownItem[];
  /**
   * Microsoft Graph Channel or Team
   *
   * @type {(MicrosoftGraph.Channel | MicrosoftGraph.Team)}
   * @memberof DropdownItem
   */
  item: MicrosoftGraph.Channel | Team;
}

/**
 * Drop down menu item state
 *
 * @interface DropdownItemState
 */
interface ChannelPickerItemState {
  /**
   * Microsoft Graph Channel or Team
   *
   * @type {(MicrosoftGraph.Channel | MicrosoftGraph.Team)}
   * @memberof ChannelPickerItemState
   */
  item: MicrosoftGraph.Channel | Team;
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
   * @type {ChannelPickerItemState[]}
   * @memberof DropdownItemState
   */
  channels?: ChannelPickerItemState[];
  /**
   * if Item has parent item (team)
   *
   * @type {ChannelPickerItemState}
   * @memberof DropdownItemState
   */
  parent: ChannelPickerItemState;
}

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
  public get selectedItem(): SelectedChannel {
    if (this._selectedItemState) {
      return { channel: this._selectedItemState.item, team: this._selectedItemState.parent.item };
    } else {
      return null;
    }
  }

  private selected: ChannelPickerItemState[] = [];

  // User input in search
  private _userInput: string = '';

  private _isHovered = false;

  private _isFocused = false;

  // Properties

  private _selectedItemState: ChannelPickerItemState;
  private _items: DropdownItem[] = [];
  private _treeViewState: ChannelPickerItemState[] = [];

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
   * @memberof MgtTeamsChannelPicker
   */
  public connectedCallback() {
    super.connectedCallback();
    window.addEventListener('click', this.handleWindowClick);
  }

  /**
   * Invoked each time the custom element is disconnected from the document's DOM
   *
   * @memberof MgtTeamsChannelPicker
   */
  public disconnectedCallback() {
    window.removeEventListener('click', this.handleWindowClick);
    super.disconnectedCallback();
  }

  /**
   * Queries the microsoft graph channel and adds them to graph
   *
   * @param {string} channelId MicrosoftGraph.Channel.id
   * @returns {Promise<void>}
   * @memberof MgtTeamsChannelPicker
   */
  public async selectChannelById(channelId: string): Promise<void> {
    // since the component normally handles loading on hover, forces the load for items
    if (this.items.length === 0) {
      this.loadTeams();
    }

    const provider = Providers.globalProvider;
    if (provider && provider.state === ProviderState.SignedIn) {
      for (const item of this._treeViewState) {
        for (const channel of item.channels) {
          if (channel.item.id === channelId) {
            this.selected = [channel];
            this._selectedItemState = channel;
            this.fireCustomEvent('selectionChanged', this.selected);
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
      this.renderTemplate('default', { teams: this.items }) ||
      html`
        <div
          class="teams-channel-picker"
          @mouseover=${this.handleHover}
          @blur=${this.lostFocus}
          @mouseout=${this.mouseLeft}
        >
          <div
            class="teams-channel-picker-input ${this._isHovered ? 'hovered' : ''} ${this._isFocused ? 'focused' : ''}"
            @click=${this.gainedFocus}
          >
            <div class="search-icon ${this._isFocused && this.selected.length === 0 ? 'focused' : ''}">
              ${this.selected.length === 0 ? getSvg(SvgIcon.Search, '#252424') : ''}
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
   * @param {ChannelPickerItemState[]} items
   * @param {number} [level=0]
   * @returns
   * @memberof MgtTeamsChannelPicker
   */
  public renderDropdown(items: ChannelPickerItemState[], level: number = 0) {
    let content: any;
    if (this.items) {
      if (this.isLoading) {
        content = this.renderTemplate('loading', null, 'loading') || this.renderLoadingMessage();
      } else if (!this.isLoading && this._treeViewState.length === 0 && this._userInput.length) {
        content = this.renderTemplate('error', null, 'error') || this.renderErrorMessage();
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
   * @param {ChannelPickerItemState} itemState
   * @returns
   * @memberof MgtTeamsChannelPicker
   */
  public renderItem(itemState: ChannelPickerItemState) {
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
   * @param {ChannelPickerItemState} item
   * @memberof MgtTeamsChannelPicker
   */
  public handleItemClick(item: ChannelPickerItemState) {
    if (item.channels) {
      item.isExpanded = !item.isExpanded;
    } else if (this._selectedItemState !== item) {
      this._selectedItemState = item;
      this.addChannel(item);
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
      }, 400);
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

    if (this.selected && this.selected.length) {
      inputClass = 'input-search';

      channelList = html`
        <li class="selected-team">
          <div class="selected-team-name">${this.selected[0].item.displayName}</div>
          <div class="arrow">${getSvg(SvgIcon.TeamSeparator, '#B3B0AD')}</div>
          ${this.selected[0].parent.item.displayName}
          <div class="CloseIcon" @click="${() => this.removeTeam()}">
            
          </div>
        </li>
        <div class="SearchIcon">
          ${this._isFocused ? getSvg(SvgIcon.Search, '#252424') : ''}
        </div>
      `;
    }
    return html`
      <div class="channel-chosen-list">
        ${channelList}
        <div class="${inputClass}">
          <input
            id="teams-channel-picker-input"
            class="team-chosen-input ${this._isFocused || this._isHovered ? 'focused' : ''}"
            type="text"
            placeholder="${this.selected.length > 0 ? '' : 'Select a channel '} "
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
    parent: ChannelPickerItemState = null
  ): ChannelPickerItemState[] {
    const treeView: ChannelPickerItemState[] = [];
    filterString = filterString.toLowerCase();

    for (const state of tree) {
      let stateItem: ChannelPickerItemState;

      if (filterString.length === 0 || state.item.displayName.toLowerCase().includes(filterString)) {
        stateItem = { item: state.item, parent };
        if (state.channels) {
          stateItem.channels = this.generateTreeViewState(state.channels, '', stateItem);
          stateItem.isExpanded = filterString.length > 0;
        }
      } else if (state.channels) {
        const newStateItem = { item: state.item, parent };
        const channels = this.generateTreeViewState(state.channels, filterString, newStateItem);
        if (channels.length > 0) {
          stateItem = newStateItem;
          stateItem.channels = channels;
          stateItem.isExpanded = true;
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
      this._isHovered = false;
      this.requestUpdate();
    }
  }

  private handleHover() {
    if (this.items.length === 0) {
      this.loadTeams();
    }
    this._isHovered = true;
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
      case 9: // tab
        if (currentFocusedItem && currentFocusedItem.channels) {
          // focus item is a Team
          currentFocusedItem.isExpanded = !currentFocusedItem.isExpanded;
          this.resetFocusState();
          event.preventDefault();
        } else if (currentFocusedItem && !currentFocusedItem.channels) {
          this._selectedItemState = currentFocusedItem;
          this.addChannel(currentFocusedItem);

          // refocus to new textbox on initial selection
          this._isFocused = true;
          this.resetFocusState();
          this._focusedIndex = -1;
          event.preventDefault();
        }
        break;
    }
  }

  private async loadTeams() {
    const provider = Providers.globalProvider;
    let teams;
    if (provider) {
      if (provider.state === ProviderState.SignedIn) {
        const graph = provider.graph.forComponent(this);

        teams = await getAllMyTeams(graph);

        this.isLoading = true;
        const batch = provider.graph.createBatch();

        for (const [i, team] of teams.entries()) {
          batch.get(`${i}`, `teams/${team.id}/channels`, ['user.read.all']);
        }
        const response = await batch.execute();
        this.items = teams.map(t => {
          return {
            item: t
          };
        });
        for (const [i] of teams.entries()) {
          this.items[i].channels = response[i].value.map(c => {
            return {
              item: c
            };
          });
        }
      }
    }
    this.items = this.items;
    this.resetFocusState();
    this.isLoading = false;
  }

  private removeTeam() {
    this.selected = [];
    this._userInput = '';

    this.fireCustomEvent('selectionChanged', this.selected);
    this.lostFocus();
    this.resetFocusState();
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
    this._isFocused = true;
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
    this._isFocused = false;
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
      channels.last = channel.displayName;
    }

    return html`
      <div class="channel-display">
        <div class="showing">
          <span class="channel-name-text">${channels.first}</span
          ><span class="channel-name-text highlight-search-text">${channels.highlight}</span
          ><span class="channel-name-text">${channels.last}</span>
        </div>
      </div>
    `;
  }

  private addChannel(item: ChannelPickerItemState) {
    this.selected = [item];
    this._userInput = '';
    this.requestUpdate();
    this.lostFocus();
    this.resetFocusState();
    this.fireCustomEvent('selectionChanged', this.selected);
  }
}
