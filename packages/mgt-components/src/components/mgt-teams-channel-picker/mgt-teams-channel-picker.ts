/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { html, TemplateResult } from 'lit';
import { state } from 'lit/decorators.js';
import { classMap } from 'lit/directives/class-map.js';
import {
  Providers,
  ProviderState,
  MgtTemplatedComponent,
  BetaGraph,
  customElement,
  mgtHtml
} from '@microsoft/mgt-element';
import '../../styles/style-helper';
import '../sub-components/mgt-spinner/mgt-spinner';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';
import { debounce } from '../../utils/Utils';
import { styles } from './mgt-teams-channel-picker-css';
import { getAllMyTeams, getTeamsPhotosforPhotoIds } from './mgt-teams-channel-picker.graph';
import { strings } from './strings';
import { repeat } from 'lit/directives/repeat.js';
import { registerFluentComponents } from '../../utils/FluentComponents';
import {
  fluentBreadcrumb,
  fluentBreadcrumbItem,
  fluentTreeView,
  fluentTreeItem,
  fluentCard,
  fluentTextField
} from '@fluentui/web-components';

registerFluentComponents(
  fluentBreadcrumb,
  fluentBreadcrumbItem,
  fluentCard,
  fluentTreeView,
  fluentTreeItem,
  fluentTextField
);

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
 * Configuration object for the TeamsChannelPicker component
 *
 * @export
 * @interface MgtTeamsChannelPickerConfig
 */
export interface MgtTeamsChannelPickerConfig {
  /**
   * Sets or gets whether the teams channel picker component should use
   * the Teams based scopes instead of the User and Group based scopes
   *
   * @type {boolean}
   */
  useTeamsBasedScopes: boolean;
}

/**
 * Web component used to select channels from a User's Microsoft Teams profile
 *
 *
 * @class MgtTeamsChannelPicker
 * @extends {MgtTemplatedComponent}
 *
 * @fires {CustomEvent<SelectedChannel | null>} selectionChanged - Fired when the selection changes
 *
 * @cssprop --color - {font} Default font color
 *
 * @cssprop --channel-picker-input-border - {String} Input section entire border
 * @cssprop --channel-picker-input-border-top - {String} Input section border top only
 * @cssprop --channel-picker-input-border-right - {String} Input section border right only
 * @cssprop --channel-picker-input-border-bottom - {String} Input section border bottom only
 * @cssprop --channel-picker-input-border-left - {String} Input section border left only
 * @cssprop --channel-picker-input-border-color - {Color} Input border color
 * @cssprop --channel-picker-input-background-color - {Color} Input section background color
 * @cssprop --channel-picker-input-background-color-hover - {Color} Input background hover color
 * @cssprop --channel-picker-input-background-color-focus - {Color} Input background focus color
 *
 * @cssprop --channel-picker-dropdown-background-color - {Color} Background color of dropdown area
 * @cssprop --channel-picker-dropdown-item-text-color - {Color} Text color of the dropdown text.
 * @cssprop --channel-picker-dropdown-item-background-color-hover - {Color} Background color of channel or team during hover
 * @cssprop --channel-picker-dropdown-item-text-color-selected - {Color} Text color of channel or team during after selection
 *
 * @cssprop --channel-picker-arrow-fill - {Color} Color of arrow svg
 * @cssprop --placeholder-color-focus - {Color} Color of placeholder text during focus state
 * @cssprop --placeholder-color - {Color} Color of placeholder text
 *
 * @cssprop --channel-picker-search-icon-color - {Color} the search icon color.
 * @cssprop --channel-picker-down-chevron-color - {Color} the down chevron icon color.
 * @cssprop --channel-picker-up-chevron-color - {Color} the up chevron icon color.
 * @cssprop --channel-picker-close-icon-color - {Color} the close icon color.
 *
 */
@customElement('teams-channel-picker')
export class MgtTeamsChannelPicker extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * user the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * Strings for localization
   *
   * @readonly
   * @protected
   * @memberof MgtTeamsChannelPicker
   */
  protected get strings() {
    return strings;
  }

  /**
   * Global Configuration object for all
   * teams channel picker components
   *
   * @static
   * @type {MgtTeamsChannelPickerConfig}
   * @memberof MgtTeamsChannelPicker
   */
  public static get config(): MgtTeamsChannelPickerConfig {
    return this._config;
  }

  private static _config = {
    useTeamsBasedScopes: false
  };
  private teamsPhotos = {};

  /**
   * Gets Selected item to be used
   *
   * @readonly
   * @type {SelectedChannel}
   * @memberof MgtTeamsChannelPicker
   */
  public get selectedItem(): SelectedChannel | null {
    if (this._selectedItemState) {
      return { channel: this._selectedItemState.item, team: this._selectedItemState.parent.item };
    } else {
      return null;
    }
  }

  /**
   * Get the scopes required for teams channel picker
   *
   * @static
   * @return {*}  {string[]}
   * @memberof MgtTeamsChannelPicker
   */
  public static get requiredScopes(): string[] {
    if (this.config.useTeamsBasedScopes) {
      return ['team.readbasic.all', 'channel.readbasic.all'];
    } else {
      return ['user.read.all', 'group.read.all'];
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
  private get _input(): HTMLInputElement {
    const wrapper = this.renderRoot.querySelector<HTMLElement>('fluent-text-field');
    const input = wrapper.shadowRoot.querySelector<HTMLInputElement>('input');
    return input;
  }
  private _inputValue = '';

  @state() private _selectedItemState: ChannelPickerItemState;
  private _items: DropdownItem[];
  private _treeViewState: ChannelPickerItemState[] = [];
  private _focusList: ChannelPickerItemState[] = [];

  // focus state
  private debouncedSearch: () => void;

  // determines loading state
  @state() private _isDropdownVisible: boolean;
  @state() private _isFocused: boolean;

  constructor() {
    super();
    this.addEventListener('focus', _ => this.loadTeamsIfNotLoaded());
    this.addEventListener('mouseover', _ => this.loadTeamsIfNotLoaded());
    this.addEventListener('blur', _ => this.lostFocus());
    this.clearState();
  }

  /**
   * Invoked each time the custom element is appended into a document-connected element
   *
   * @memberof MgtTeamsChannelPicker
   */
  public connectedCallback() {
    super.connectedCallback();
    window.addEventListener('click', this.handleWindowClick);

    const ownerDocument = this.renderRoot.ownerDocument;
    if (ownerDocument) {
      ownerDocument.documentElement.setAttribute('dir', this.direction);
    }
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
            item.isExpanded = true;
            this.selectChannel(channel);
            this.markSelectedChannelInDropdown(channelId);
            return true;
          }
        }
      }
    }
    return false;
  }

  /**
   * Marks a channel selected by ID as selected in the dropdown menu.
   * It ensures the parent team is set to as expanded to show the channel.
   *
   * @param channelId ID string of the selected channel
   */
  private markSelectedChannelInDropdown(channelId: string) {
    const treeItem = this.renderRoot.querySelector(`[id='${channelId}']`);
    if (treeItem) {
      treeItem.setAttribute('selected', 'true');
      if (treeItem.parentElement) {
        treeItem.parentElement.setAttribute('expanded', 'true');
      }
    }
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return a lit-html TemplateResult.
   * Setting properties inside this method will not trigger the element to update.
   *
   * @returns
   * @memberof MgtTeamsChannelPicker
   */
  public render() {
    const dropdownClasses = {
      dropdown: true,
      visible: this._isDropdownVisible
    };

    return (
      this.renderTemplate('default', { teams: this.items }) ||
      html`
        <div class="container" @blur=${this.lostFocus}>
          <fluent-text-field
            appearance="outline"
            id="teams-channel-picker-input"
            aria-label="Select a channel"
            placeholder="${!!this._selectedItemState ? '' : this.strings.inputPlaceholderText} "
            label="teams-channel-picker-input"
            @click=${this.gainedFocus}
            @keyup=${(e: KeyboardEvent) => this.handleInputChanged(e)}>
              <div slot="start" style="width: max-content;">${this.renderSelected()}</div>
              <div slot="end">${this.renderChevrons()}${this.renderCloseButton()}</div>
          </fluent-text-field>
          <fluent-card class=${classMap(dropdownClasses)}>
            ${this.renderDropdown()}
          </fluent-card>
        </div>`
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
      return this.renderSearchIcon();
    }
    let icon: TemplateResult;
    if (this._selectedItemState.parent.channels) {
      // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-unsafe-assignment
      const src = this.teamsPhotos[this._selectedItemState.parent.item.id]?.photo;
      icon = html`<img
        class="team-photo"
        alt="${this._selectedItemState.parent.item.displayName}"
        src=${src} />`;
    }

    const parentName = this._selectedItemState?.parent?.item?.displayName.trim();
    const channelName = this._selectedItemState?.item?.displayName.trim();

    return html`
      <fluent-breadcrumb title=${this._selectedItemState.item.displayName}>
        <fluent-breadcrumb-item>
          <span slot="start">${icon}</span>
          <span class="team-parent-name">${parentName}</span>
          <span slot="separator" class="arrow">${getSvg(SvgIcon.TeamSeparator, '#000000')}</span>
        </fluent-breadcrumb-item>
        <fluent-breadcrumb-item>${channelName}</fluent-breadcrumb-item>
      </fluent-breadcrumb>`;
  }

  /**
   * Clears the state of the component
   *
   * @protected
   * @memberof MgtTeamsChannelPicker
   */
  protected clearState(): void {
    this._items = [];
    this._inputValue = '';
    this._treeViewState = [];
    this._focusList = [];
    this._isDropdownVisible = false;
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
   * Renders close button
   *
   * @protected
   * @returns
   * @memberof MgtTeamsChannelPicker
   */
  protected renderCloseButton() {
    return html`
      <div
        class="close-icon"
        style="display:none"
        @click="${() => this.removeSelectedChannel(null)}">
        <svg width="8" height="8" viewBox="0 0 8 8" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M0.0885911 0.215694L0.146447 0.146447C0.320013 -0.0271197 0.589437 -0.046405 0.784306 0.0885911L0.853553 0.146447L4 3.293L7.14645 0.146447C7.34171 -0.0488154 7.65829 -0.0488154 7.85355 0.146447C8.04882 0.341709 8.04882 0.658291 7.85355 0.853553L4.707 4L7.85355 7.14645C8.02712 7.32001 8.0464 7.58944 7.91141 7.78431L7.85355 7.85355C7.67999 8.02712 7.41056 8.0464 7.21569 7.91141L7.14645 7.85355L4 4.707L0.853553 7.85355C0.658291 8.04882 0.341709 8.04882 0.146447 7.85355C-0.0488154 7.65829 -0.0488154 7.34171 0.146447 7.14645L3.293 4L0.146447 0.853553C-0.0271197 0.679987 -0.046405 0.410563 0.0885911 0.215694L0.146447 0.146447L0.0885911 0.215694Z" fill="#212121"/>
        </svg>
      </div>
    `;
  }

  /**
   * Displays the close button after selecting a channel.
   */
  protected showCloseIcon() {
    const downChevron = this.renderRoot.querySelector<HTMLElement>('.down-chevron');
    const upChevron = this.renderRoot.querySelector<HTMLElement>('.up-chevron');
    const closeIcon = this.renderRoot.querySelector<HTMLElement>('.close-icon');
    if (downChevron) {
      downChevron.style.display = 'none';
    }
    if (upChevron) {
      upChevron.style.display = 'none';
    }

    if (closeIcon) {
      closeIcon.style.display = null;
    }
  }

  /**
   * Renders down chevron icon
   *
   * @protected
   * @returns
   * @memberof MgtTeamsChannelPicker
   */
  protected renderDownChevron() {
    return html`
      <div class="down-chevron" @click=${this.gainedFocus}>
        <svg width="12" height="12" viewBox="0 0 12 12" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M2.21967 4.46967C2.51256 4.17678 2.98744 4.17678 3.28033 4.46967L6 7.18934L8.71967 4.46967C9.01256 4.17678 9.48744 4.17678 9.78033 4.46967C10.0732 4.76256 10.0732 5.23744 9.78033 5.53033L6.53033 8.78033C6.23744 9.07322 5.76256 9.07322 5.46967 8.78033L2.21967 5.53033C1.92678 5.23744 1.92678 4.76256 2.21967 4.46967Z" fill="#212121" />
        </svg>
      </div>`;
  }

  /**
   * Renders up chevron icon
   *
   * @protected
   * @returns
   * @memberof MgtTeamsChannelPicker
   */
  protected renderUpChevron() {
    return html`
      <div style="display:none" class="up-chevron" @click=${(e: Event) => this.handleUpChevronClick(e)}>
        <svg width="12" height="12" viewBox="0 0 12 12" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M2.21967 7.53033C2.51256 7.82322 2.98744 7.82322 3.28033 7.53033L6 4.81066L8.71967 7.53033C9.01256 7.82322 9.48744 7.82322 9.78033 7.53033C10.0732 7.23744 10.0732 6.76256 9.78033 6.46967L6.53033 3.21967C6.23744 2.92678 5.76256 2.92678 5.46967 3.21967L2.21967 6.46967C1.92678 6.76256 1.92678 7.23744 2.21967 7.53033Z" fill="#212121" />
        </svg>
      </div>`;
  }

  /**
   * Renders both chevrons
   */
  private renderChevrons() {
    return html`${this.renderUpChevron()}${this.renderDownChevron()}`;
  }

  /**
   * Renders dropdown content
   *
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
   * @returns
   * @memberof MgtTeamsChannelPicker
   */
  protected renderDropdownList(items: ChannelPickerItemState[]) {
    if (items && items.length > 0) {
      let icon: TemplateResult = null;

      return html`
        <fluent-tree-view
          class="tree-view"
          dir=${this.direction}>
          ${repeat(
            items,
            (itemObj: ChannelPickerItemState) => itemObj?.item,
            (obj: ChannelPickerItemState) => {
              if (obj.channels) {
                icon = html`<img
                  class="team-photo"
                  alt="${obj.item.displayName}"
                  src=${
                    // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access
                    this.teamsPhotos[obj.item.id]?.photo
                  } />`;
              }
              return html`
                <fluent-tree-item
                  ?expanded=${obj?.isExpanded}
                  @click=${(e: Event) => this.handleTeamTreeItemClick(e)}>
                    ${icon}${obj.item.displayName}
                    ${repeat(
                      obj?.channels,
                      (channels: ChannelPickerItemState) => channels.item,
                      (channel: ChannelPickerItemState) => {
                        return this.renderItem(channel);
                      }
                    )}
                </fluent-tree-item>`;
            }
          )}
        </fluent-tree-view>`;
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
    return html`
      <fluent-tree-item
        id=${itemState?.item?.id}
        @keydown=${(e: KeyboardEvent) => this.onUserKeyDown(e, itemState)}
        @click=${() => this.handleItemClick(itemState)}>
          ${itemState?.item.displayName}
      </fluent-tree-item>`;
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
          <div
            label="search-error-text"
            aria-label="We didn't find any matches."
            class="search-error-text">
            ${this.strings.noResultsFound}
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
      mgtHtml`
        <div class="message-parent">
          <mgt-spinner></mgt-spinner>
          <div label="loading-text" aria-label="loading" class="loading-text">
            ${this.strings.loadingMessage}
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

      // make sure we have the needed scopes
      if (!(await provider.getAccessTokenForScopes(...MgtTeamsChannelPicker.requiredScopes))) {
        return;
      }

      teams = await getAllMyTeams(graph);
      teams = teams.filter(t => !t.isArchived);

      const teamsIds = teams.map(t => t.id);

      const beta = BetaGraph.fromGraph(graph);

      this.teamsPhotos = await getTeamsPhotosforPhotoIds(beta, teamsIds);

      const batch = graph.createBatch();
      const scopes = ['team.readbasic.all'];

      for (const team of teams) {
        batch.get(team.id, `teams/${team.id}/channels`, scopes);
      }

      const responses = await batch.executeAll();

      for (const team of teams) {
        const response = responses.get(team.id);

        // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access
        if (response?.content?.value) {
          // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-unsafe-call
          team.channels = response.content.value.map((c: MicrosoftGraph.Team) => {
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

  /**
   * Handles operations that are performed on the DOM when you remove a
   * channel. For example on clicking the X button.
   *
   * @param item a selected channel item
   */
  private removeSelectedChannel(item: ChannelPickerItemState) {
    this.selectChannel(item);
    const treeItems = this.renderRoot.querySelectorAll('fluent-tree-item');
    if (treeItems) {
      treeItems.forEach((treeItem: HTMLElement) => {
        treeItem.removeAttribute('expanded');
        treeItem.removeAttribute('selected');
      });
    }
  }

  private handleItemClick(item: ChannelPickerItemState) {
    if (item.channels) {
      item.isExpanded = !item.isExpanded;
    } else {
      this.selectChannel(item);
      this.lostFocus();
    }
  }

  private handleTeamTreeItemClick(event: Event) {
    event.preventDefault();
    event.stopImmediatePropagation();
    const element = event.target as HTMLElement;
    if (element) {
      const expanded = element.getAttribute('expanded');

      if (!!expanded) {
        element.removeAttribute('expanded');
      } else {
        element.setAttribute('expanded', 'true');
      }
      element.removeAttribute('selected');
      const hasId = element.getAttribute('id');
      if (hasId) {
        element.setAttribute('selected', 'true');
      }
    }
  }

  private handleInputChanged(e: KeyboardEvent) {
    const target = e.target as HTMLInputElement;
    if (this._inputValue !== target?.value) {
      this._inputValue = target?.value;
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

  private onUserKeyDown(e: KeyboardEvent, item?: ChannelPickerItemState) {
    const key = e.code;
    switch (key) {
      case 'Enter':
        this.selectChannel(item);
        this.resetFocusState();
        this.lostFocus();
        e.preventDefault();
        break;
      case 'Backspace':
        if (this._inputValue.length === 0 && this._selectedItemState) {
          this.selectChannel(null);
          this.resetFocusState();
          e.preventDefault();
        }
        break;
    }
  }

  private filterList() {
    if (this.items) {
      this._treeViewState = this.generateTreeViewState(this.items, this._inputValue);
      this.resetFocusState();
    }
  }

  private generateTreeViewState(
    tree: DropdownItem[],
    filterString = '',
    parent: ChannelPickerItemState = null
  ): ChannelPickerItemState[] {
    const treeView: ChannelPickerItemState[] = [];
    filterString = filterString.toLowerCase();

    if (tree) {
      for (const item of tree) {
        let stateItem: ChannelPickerItemState;

        if (filterString.length === 0 || item.item.displayName.toLowerCase().includes(filterString)) {
          stateItem = { item: item.item, parent };
          if (item.channels) {
            stateItem.channels = this.generateTreeViewState(item.channels, '', stateItem);
            stateItem.isExpanded = filterString.length > 0;
          }
        } else if (item.channels) {
          const newStateItem = { item: item.item, parent };
          const channels = this.generateTreeViewState(item.channels, filterString, newStateItem);
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
  private generateFocusList(items: ChannelPickerItemState[]): ChannelPickerItemState[] {
    if (!items || items.length === 0) {
      return [];
    }

    let array: ChannelPickerItemState[] = [];

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
      void this.requestStateUpdate();
    }
  }

  private handleWindowClick = (e: MouseEvent) => {
    if (e.target !== this) {
      this.lostFocus();
    }
  };

  private gainedFocus = () => {
    this._isFocused = true;
    const input = this._input;
    if (input) {
      input.focus();
    }

    this._isDropdownVisible = true;
    this.toggleChevron();
    this.resetFocusState();
  };

  private lostFocus = () => {
    const input = this._input;
    if (input) {
      input.value = this._inputValue = '';
      input.textContent = '';
    }

    this._isFocused = false;
    this._isDropdownVisible = false;
    this.filterList();
    this.toggleChevron();
    this.requestUpdate();

    if (this._selectedItemState !== undefined) {
      this.showCloseIcon();
    }
  };

  private selectChannel(item: ChannelPickerItemState) {
    if (item && this._selectedItemState !== item) {
      this._input.setAttribute('disabled', 'true');
    } else {
      this._input.removeAttribute('disabled');
    }
    this._selectedItemState = item;
    this.lostFocus();
    this.fireCustomEvent('selectionChanged', this.selectedItem);
  }

  /**
   * Hides the close icon.
   */
  private hideCloseIcon() {
    const closeIcon = this.renderRoot.querySelector<HTMLElement>('.close-icon');
    if (closeIcon) {
      closeIcon.style.display = 'none';
    }
  }

  /**
   * Toggles the up and down chevron depending on the dropdown
   * visibility.
   */
  private toggleChevron() {
    const downChevron = this.renderRoot.querySelector<HTMLElement>('.down-chevron');
    const upChevron = this.renderRoot.querySelector<HTMLElement>('.up-chevron');
    if (this._isDropdownVisible) {
      if (downChevron) {
        downChevron.style.display = 'none';
      }
      if (upChevron) {
        upChevron.style.display = null;
      }
    } else {
      if (downChevron) {
        downChevron.style.display = null;
        this.hideCloseIcon();
      }
      if (upChevron) {
        upChevron.style.display = 'none';
      }
    }
    this.hideCloseIcon();
  }

  private handleUpChevronClick(e: Event) {
    e.stopPropagation();
    this.lostFocus();
  }
}
