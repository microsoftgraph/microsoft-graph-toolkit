/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property, TemplateResult } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { Providers, ProviderState } from '@microsoft/mgt-element';
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
 * @fires selectionChanged - Fired when the selection changes
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
 * @cssprop --input-focus-color - {Color} Input text focus color
 *
 * @cssprop --dropdown-background-color - {Color} Background color of dropdown area
 * @cssprop --dropdown-item-hover-background - {Color} Background color of channel or team during hover
 * @cssprop --dropdown-item-selected-background - {Color} Background color of selected channel
 *
 * @cssprop --arrow-fill - {Color} Color of arrow svg
 * @cssprop --placeholder-focus-color - {Color} Color of placeholder text during focus state
 * @cssprop --placeholder-default-color - {Color} Color of placeholder text
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

  /**
   * Gets Selected item to be used
   *
   * @readonly
   * @type {SelectedChannel}
   * @memberof MgtTeamsChannelPicker
   */
  public get selectedItem(): SelectedChannel {
    if (this._selectedItemState) {
      return { channel: this._selectedItemState.item, team: this._selectedItemState.parent.item };
    } else {
      return null;
    }
  }

  private set items(value) {
    if (this._items === value) {
      return;
    }
    this._items = value;
    this._treeViewState = value ? this.generateTreeViewState(value) : [];
    this.resetFocusState();
  }
  private get items(): DropdownItem[] {
    return this._items;
  }

  // User input in search
  private get _input(): HTMLElement {
    return this.renderRoot.querySelector('.team-chosen-input');
  }
  private _inputValue: string = '';

  private _isFocused = false;

  private _selectedItemState: ChannelPickerItemState;
  private _items: DropdownItem[];
  private _treeViewState: ChannelPickerItemState[] = [];

  // focus state
  private _focusList: ChannelPickerItemState[] = [];
  private _focusedIndex: number = -1;
  private debouncedSearch;

  // determines loading state
  @property({ attribute: false }) private _isDropdownVisible;

  constructor() {
    super();
    this.handleWindowClick = this.handleWindowClick.bind(this);
    this.addEventListener('keydown', e => this.onUserKeyDown(e));
    this.addEventListener('focus', _ => this.loadTeamsIfNotLoaded());
    this.addEventListener('mouseover', _ => this.loadTeamsIfNotLoaded());
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
   * selects a channel by looking up the id in the Graph
   *
   * @param {string} channelId MicrosoftGraph.Channel.id
   * @returns {Promise<return>} A promise that will resolve to true if channel was selected
   * @memberof MgtTeamsChannelPicker
   */
  public async selectChannelById(channelId: string): Promise<boolean> {
    const provider = Providers.globalProvider;
    if (provider && provider.state === ProviderState.SignedIn) {
      // since the component normally handles loading on hover, forces the load for items
      if (!this.items) {
        await this.requestStateUpdate();
      }

      for (const item of this._treeViewState) {
        for (const channel of item.channels) {
          if (channel.item.id === channelId) {
            this.selectChannel(channel);
            return true;
          }
        }
      }
    }
    return false;
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return a lit-html TemplateResult.
   * Setting properties inside this method will not trigger the element to update.
   * @returns
   * @memberof MgtTeamsChannelPicker
   */
  public render() {
    const inputClasses = {
      focused: this._isFocused,
      'input-wrapper': true
    };

    const iconClasses = {
      focused: this._isFocused && !!this._selectedItemState,
      'search-icon': true
    };

    const dropdownClasses = {
      dropdown: true,
      visible: this._isDropdownVisible
    };

    return (
      this.renderTemplate('default', { teams: this.items }) ||
      html`
        <div class="root" @blur=${this.lostFocus}>
          <div class=${classMap(inputClasses)} @click=${this.gainedFocus}>
            ${this.renderSelected()}
            <div class="search-wrapper">${this.renderSearchIcon()} ${this.renderInput()}</div>
          </div>
          ${this.renderCloseButton()}
          <div class=${classMap(dropdownClasses)}>${this.renderDropdown()}</div>
        </div>
      `
    );
  }

  /**
   * Renders selected channel
   *
   * @protected
   * @returns
   * @memberof MgtTeamsChannelPicker
   */
  protected renderSelected() {
    if (!this._selectedItemState) {
      return html``;
    }

    return html`
      <li class="selected-team" title=${this._selectedItemState.item.displayName}>
        <div class="selected-team-name">${this._selectedItemState.parent.item.displayName}</div>
        <div class="arrow">${getSvg(SvgIcon.TeamSeparator, '#B3B0AD')}</div>
        ${this._selectedItemState.item.displayName}
      </li>
    `;
  }

  /**
   * Renders search icon
   *
   * @protected
   * @returns
   * @memberof MgtTeamsChannelPicker
   */
  protected renderSearchIcon() {
    return html`
      <div class="search-icon">
        ${getSvg(SvgIcon.Search, '#252424')}
      </div>
    `;
  }

  /**
   * Renders input field
   *
   * @protected
   * @returns
   * @memberof MgtTeamsChannelPicker
   */
  protected renderInput() {
    const rootClasses = {
      'input-search': !!this._selectedItemState,
      'input-search-start': !this._selectedItemState
    };

    return html`
      <div class="${classMap(rootClasses)}">
        <span
          id="teams-channel-picker-input"
          class="team-chosen-input"
          type="text"
          label="teams-channel-picker-input"
          aria-label="Select a channel"
          data-placeholder="${!!this._selectedItemState ? '' : 'Select a channel '} "
          role="input"
          @keyup=${e => this.handleInputChanged(e)}
          contenteditable
        ></span>
      </div>
    `;
  }

  /**
   * Renders close button
   *
   * @protected
   * @returns
   * @memberof MgtTeamsChannelPicker
   */
  protected renderCloseButton() {
    return html`
      <div class="close-icon" @click="${() => this.selectChannel(null)}">
        
      </div>
    `;
  }

  /**
   * Renders dropdown content
   *
   * @param {ChannelPickerItemState[]} items
   * @param {number} [level=0]
   * @returns
   * @memberof MgtTeamsChannelPicker
   */
  protected renderDropdown() {
    if (this.isLoadingState || !this._treeViewState) {
      return this.renderLoading();
    }

    if (this._treeViewState) {
      if (!this.isLoadingState && this._treeViewState.length === 0 && this._inputValue.length > 0) {
        return this.renderError();
      }

      return this.renderDropdownList(this._treeViewState);
    }

    return html``;
  }

  /**
   * Renders the dropdown list recursively
   *
   * @protected
   * @param {ChannelPickerItemState[]} items
   * @param {number} [level=0]
   * @returns
   * @memberof MgtTeamsChannelPicker
   */
  protected renderDropdownList(items: ChannelPickerItemState[], level: number = 0) {
    if (items && items.length) {
      return items.map((treeItem, index) => {
        const isLeaf = !treeItem.channels;
        const renderChannels = !isLeaf && treeItem.isExpanded;

        return html`
          ${this.renderItem(treeItem)}
          ${renderChannels ? this.renderDropdownList(treeItem.channels, level + 1) : html``}
        `;
      });
    }

    return null;
  }

  /**
   * Renders each Channel or Team
   *
   * @param {ChannelPickerItemState} itemState
   * @returns
   * @memberof MgtTeamsChannelPicker
   */
  protected renderItem(itemState: ChannelPickerItemState) {
    let icon: TemplateResult = null;

    if (itemState.channels) {
      // must be team with channels
      icon = itemState.isExpanded ? getSvg(SvgIcon.ArrowDown, '#252424') : getSvg(SvgIcon.ArrowRight, '#252424');
    }

    let isSelected = false;
    if (this.selectedItem) {
      if (this.selectedItem.channel === itemState.item) {
        isSelected = true;
      }
    }

    const classes = {
      focused: this._focusList[this._focusedIndex] === itemState,
      item: true,
      'list-team': itemState.channels ? true : false,
      selected: isSelected
    };

    const dropDown = this.renderRoot.querySelector('.dropdown');

    if (dropDown.children[this._focusedIndex]) {
      dropDown.children[this._focusedIndex].scrollIntoView(false);
    }

    return html`
      <div @click=${() => this.handleItemClick(itemState)} class="${classMap(classes)}">
        <div class="arrow">
          ${icon}
        </div>
        ${itemState.channels ? itemState.item.displayName : this.renderHighlightedText(itemState.item)}
      </div>
    `;
  }

  /**
   * Renders the channel with the query text higlighted
   *
   * @protected
   * @param {*} channel
   * @returns
   * @memberof MgtTeamsChannelPicker
   */
  protected renderHighlightedText(channel: any) {
    // tslint:disable-next-line: prefer-const
    let channels: any = {};

    const highlightLocation = channel.displayName.toLowerCase().indexOf(this._inputValue.toLowerCase());
    if (highlightLocation !== -1) {
      // no location
      if (highlightLocation === 0) {
        // highlight is at the beginning of sentence
        channels.first = '';
        channels.highlight = channel.displayName.slice(0, this._inputValue.length);
        channels.last = channel.displayName.slice(this._inputValue.length, channel.displayName.length);
      } else if (highlightLocation === channel.displayName.length) {
        // highlight is at end of the sentence
        channels.first = channel.displayName.slice(0, highlightLocation);
        channels.highlight = channel.displayName.slice(highlightLocation, channel.displayName.length);
        channels.last = '';
      } else {
        // highlight is in middle of sentence
        channels.first = channel.displayName.slice(0, highlightLocation);
        channels.highlight = channel.displayName.slice(highlightLocation, highlightLocation + this._inputValue.length);
        channels.last = channel.displayName.slice(
          highlightLocation + this._inputValue.length,
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

  /**
   * Renders an error message when no channel or teams match the query
   *
   * @protected
   * @returns
   * @memberof MgtTeamsChannelPicker
   */
  protected renderError(): TemplateResult {
    const template = this.renderTemplate('error', null, 'error');

    return (
      template ||
      html`
        <div class="message-parent">
          <div label="search-error-text" aria-label="We didn't find any matches." class="search-error-text">
            We didn't find any matches.
          </div>
        </div>
      `
    );
  }

  /**
   * Renders loading spinner while channels are fetched from the Graph
   *
   * @protected
   * @returns
   * @memberof MgtTeamsChannelPicker
   */
  protected renderLoading(): TemplateResult {
    const template = this.renderTemplate('loading', null, 'loading');

    return (
      template ||
      html`
        <div class="message-parent">
          <div class="spinner"></div>
          <div label="loading-text" aria-label="loading" class="loading-text">
            Loading...
          </div>
        </div>
      `
    );
  }

  /**
   * Queries Microsoft Graph for Teams & respective channels then sets to items list
   *
   * @protected
   * @memberof MgtTeamsChannelPicker
   */
  protected async loadState() {
    const provider = Providers.globalProvider;
    let teams: MicrosoftGraph.Team[];
    if (provider && provider.state === ProviderState.SignedIn) {
      const graph = provider.graph.forComponent(this);

      teams = await getAllMyTeams(graph);
      teams = teams.filter(t => !t.isArchived);

      const batch = provider.graph.createBatch();

      for (const team of teams) {
        batch.get(team.id, `teams/${team.id}/channels`, ['group.read.all']);
      }

      const responses = await batch.executeAll();

      for (const team of teams) {
        const response = responses.get(team.id);

        if (response && response.content && response.content.value) {
          team.channels = response.content.value.map(c => {
            return {
              item: c
            };
          });
        }
      }

      this.items = teams.map(t => {
        return {
          channels: t.channels as DropdownItem[],
          item: t
        };
      });
    }
    this.filterList();
    this.resetFocusState();
  }

  private handleItemClick(item: ChannelPickerItemState) {
    if (item.channels) {
      item.isExpanded = !item.isExpanded;
    } else {
      this.selectChannel(item);
    }

    this._focusedIndex = -1;
    this.resetFocusState();
  }

  private handleInputChanged(e) {
    if (this._inputValue !== e.target.textContent) {
      this._inputValue = e.target.textContent;
    } else {
      return;
    }

    // shows list
    this.gainedFocus();

    if (!this.debouncedSearch) {
      this.debouncedSearch = debounce(() => {
        this.filterList();
      }, 400);
    }

    this.debouncedSearch();
  }

  private filterList() {
    if (this.items) {
      this._treeViewState = this.generateTreeViewState(this.items, this._inputValue);
      this._focusedIndex = -1;
      this.resetFocusState();
    }
  }

  private generateTreeViewState(
    tree: DropdownItem[],
    filterString: string = '',
    parent: ChannelPickerItemState = null
  ): ChannelPickerItemState[] {
    const treeView: ChannelPickerItemState[] = [];
    filterString = filterString.toLowerCase();

    if (tree) {
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

  private loadTeamsIfNotLoaded() {
    if (!this.items && !this.isLoadingState) {
      this.requestStateUpdate();
    }
  }

  private handleWindowClick(e: MouseEvent) {
    if (e.target !== this) {
      this.lostFocus();
    }
  }

  private onUserKeyDown(event: KeyboardEvent) {
    if (event.keyCode === 13) {
      // No new line
      event.preventDefault();
    }

    if (this._treeViewState.length === 0) {
      return;
    }

    const currentFocusedItem = this._focusList[this._focusedIndex];

    switch (event.keyCode) {
      case 40: // down
        this._focusedIndex = (this._focusedIndex + 1) % this._focusList.length;
        this.requestUpdate();
        event.preventDefault();
        break;
      case 38: // up
        if (this._focusedIndex === -1) {
          this._focusedIndex = this._focusList.length;
        }
        this._focusedIndex = (this._focusedIndex - 1 + this._focusList.length) % this._focusList.length;
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
        if (!currentFocusedItem) {
          this.lostFocus();
          break;
        }
      case 13: // return/enter
        if (currentFocusedItem && currentFocusedItem.channels) {
          // focus item is a Team
          currentFocusedItem.isExpanded = !currentFocusedItem.isExpanded;
          this.resetFocusState();
          event.preventDefault();
        } else if (currentFocusedItem && !currentFocusedItem.channels) {
          this.selectChannel(currentFocusedItem);

          // refocus to new textbox on initial selection
          this.resetFocusState();
          this._focusedIndex = -1;
          event.preventDefault();
        }
        break;
      case 8: // backspace
        if (this._inputValue.length === 0 && this._selectedItemState) {
          this.selectChannel(null);
          event.preventDefault();
        }
        break;
      case 27: // esc
        this.selectChannel(this._selectedItemState);
        this._focusedIndex = -1;
        this.resetFocusState();
        event.preventDefault();
        break;
    }
  }

  private gainedFocus() {
    this._isFocused = true;
    const input = this._input;
    if (input) {
      input.focus();
    }

    this._isDropdownVisible = true;
  }

  private lostFocus() {
    this._isFocused = false;
    const input = this._input;
    if (input) {
      input.textContent = this._inputValue = '';
    }

    this._isDropdownVisible = false;
    this.filterList();
  }

  private selectChannel(item: ChannelPickerItemState) {
    if (this._selectedItemState !== item) {
      this._selectedItemState = item;
      this.fireCustomEvent('selectionChanged', item ? [this.selectedItem] : []);
    }

    const input = this._input;
    if (input) {
      input.textContent = this._inputValue = '';
    }
    this.requestUpdate();
    this.lostFocus();
  }
}
