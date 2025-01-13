/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html, TemplateResult } from 'lit';
import { property, state } from 'lit/decorators.js';
import { repeat } from 'lit/directives/repeat.js';
import { TaskStatus } from '@lit/task';
import {
  findGroups,
  getGroupsForGroupIds,
  GroupType,
  getGroup,
  findGroupsFromGroupIds,
  isGroupType,
  groupTypeConverter
} from '../../graph/graph.groups';
import {
  findPeople,
  getPeople,
  PersonType,
  personTypeConverter,
  UserType,
  userTypeConverter
} from '../../graph/graph.people';
import {
  findUsers,
  findGroupMembers,
  findUsersFromGroupIds,
  getUser,
  getUsersForUserIds,
  getUsers
} from '../../graph/graph.user';
import { IDynamicPerson } from '../../graph/types';
import {
  Providers,
  ProviderState,
  MgtTemplatedTaskComponent,
  arraysAreEqual,
  IGraph,
  mgtHtml
} from '@microsoft/mgt-element';
import '../../styles/style-helper';
import { debounce, isValidEmail } from '../../utils/Utils';
import { MgtPerson, defaultPersonProperties, registerMgtPersonComponent } from '../mgt-person/mgt-person';
import { MgtFlyout, registerMgtFlyoutComponent } from '../sub-components/mgt-flyout/mgt-flyout';
import { styles } from './mgt-people-picker-css';
import { SvgIcon, getSvg } from '../../utils/SvgHelper';
import { fluentTextField, fluentCard } from '@fluentui/web-components';
import { registerFluentComponents } from '../../utils/FluentComponents';
import { strings } from './strings';
import { Person, User } from '@microsoft/microsoft-graph-types';
import { registerComponent } from '@microsoft/mgt-element';
import { registerMgtSpinnerComponent } from '../sub-components/mgt-spinner/mgt-spinner';
import { isGraphError } from '../../graph/isGraphError';
import { type PersonCardInteraction, personCardConverter } from './../PersonCardInteraction';

export { GroupType } from '../../graph/graph.groups';
export { PersonType, UserType } from '../../graph/graph.people';

export const registerMgtPeoplePickerComponent = () => {
  registerFluentComponents(fluentTextField, fluentCard);

  registerMgtFlyoutComponent();
  registerMgtPersonComponent();
  registerMgtSpinnerComponent();
  registerComponent('people-picker', MgtPeoplePicker);
};

/**
 * Web component used to search for people from the Microsoft Graph
 *
 * @export
 * @class MgtPicker
 * @extends {MgtTemplatedTaskComponent}
 *
 * @fires {CustomEvent<undefined>} updated - Fired when the component is updated
 * @fires {CustomEvent<IDynamicPerson[]>} selectionChanged - Fired when set of selected people changes
 *
 * @cssprop --people-picker-selected-option-background-color - {Color} the background color of the selected person.
 * @cssprop --people-picker-selected-option-highlight-background-color - {Color} the background color of the selected person when you select it for copy/cut.
 * @cssprop --people-picker-dropdown-background-color - {Color} the background color of the dropdown card.
 * @cssprop --people-picker-dropdown-result-background-color - {Color} the background color of the dropdown result.
 * @cssprop --people-picker-dropdown-result-hover-background-color - {Color} the background color of the dropdown result on hover.
 * @cssprop --people-picker-dropdown-result-focus-background-color - {Color} the background color of the dropdown result on focus.
 * @cssprop --people-picker-no-results-text-color - {Color} the no results found text color.
 * @cssprop --people-picker-input-background - {Color} the input background color.
 * @cssprop --people-picker-input-border-color - {Color} the input border color.
 * @cssprop --people-picker-input-hover-background - {Color} the input background color when you hover.
 * @cssprop --people-picker-input-hover-border-color - {Color} the input border color when you hover
 * @cssprop --people-picker-input-focus-background - {Color} the input background color when you focus.
 * @cssprop --people-picker-input-focus-border-color - {Color} the input border color when you focus.
 * @cssprop --people-picker-input-placeholder-focus-text-color - {Color} the placeholder text color when you focus.
 * @cssprop --people-picker-input-placeholder-hover-text-color - {Color} the placeholder text color when you hover.
 * @cssprop --people-picker-input-placeholder-text-color - {Color} the placeholder text color.
 * @cssprop --people-picker-search-icon-color - {Color} the search icon color
 * @cssprop --people-picker-remove-selected-close-icon-color - {Color} the remove selected person close icon color.
 * @cssprop --people-picker-result-person-avatar-size - {Length} the avatar size of the person in the result. Default is 40px.
 * @cssprop --people-picker-selected-person-avatar-size - {Length} the avatar size of the selected person. Default is 24px.
 * @cssprop --people-picker-font-size - {Length} the font size of the text in the people picker input. Default is 14px.
 */
export class MgtPeoplePicker extends MgtTemplatedTaskComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * user the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * The strings to be used for localizing the component.
   *
   * @readonly
   * @protected
   * @memberof MgtPeoplePicker
   */
  protected get strings() {
    return strings;
  }

  /**
   * Gets the flyout element
   *
   * @protected
   * @type {MgtFlyout}
   * @memberof MgtLogin
   */
  protected get flyout(): MgtFlyout {
    return this.renderRoot.querySelector('.flyout');
  }

  /**
   * Gets the input element
   *
   * @protected
   * @type {MgtFlyout}
   * @memberof MgtLogin
   */
  protected get input(): HTMLInputElement {
    return this.renderRoot.querySelector('fluent-text-field');
  }

  /**
   * value determining if search is filtered to a group.
   *
   * @type {string}
   */
  @property({ attribute: 'group-id', converter: value => value.trim() })
  public groupId: string;

  /**
   * array of groups for search to be filtered by.
   *
   * @type {string[]}
   */
  @property({
    attribute: 'group-ids',
    converter: value => {
      return value.split(',').map(v => v.trim());
    }
  })
  public groupIds: string[] = [];
  /**
   * Value determining if search is filtered to a group.
   * Valid options are 'any', 'person', 'group'
   *
   * @type {PersonType}
   */
  @property({
    attribute: 'type',
    converter: value => personTypeConverter(value, 'any')
  })
  public type: PersonType = 'any';

  /**
   * Type of groups to search for - requires personType to be set to "Group" or "All"
   * Valid values are 'any', 'unified', 'security', 'mailenabledsecurity', 'distribution'
   * Default is ['any'].
   *
   * @type {GroupType}
   */
  @property({
    attribute: 'group-type',
    converter: value => {
      const defaultResult: GroupType[] = ['any'];
      if (!value || value.length === 0) {
        return defaultResult;
      }

      const values = value.split(',');
      const groupTypes: GroupType[] = [];

      for (let v of values) {
        v = v.trim();
        if (isGroupType(v)) {
          groupTypes.push(groupTypeConverter(v));
        }
      }

      if (groupTypes.length === 0) {
        return defaultResult;
      }

      return groupTypes;
    }
  })
  public groupType: GroupType[] = ['any'];

  /**
   * The type of user to search for.
   * Valid options are 'any', 'user', 'contact'
   * Default is any.
   *
   * @type {UserType}
   * @memberof MgtPeoplePicker
   */
  @property({
    attribute: 'user-type',
    converter: value => userTypeConverter(value)
  })
  public userType: UserType = 'any';

  /**
   * whether the return should contain a flat list of all nested members
   *
   * @type {boolean}
   */
  @property({
    attribute: 'transitive-search',
    type: Boolean
  })
  public transitiveSearch = false;

  /**
   * containing object of IDynamicPerson.
   *
   * @type {IDynamicPerson[]}
   */
  @property({
    attribute: 'people',
    type: Object
  })
  public people: IDynamicPerson[] | null = null;

  /**
   * determining how many people to show in list.
   *
   * @type {number}
   */
  @property({
    attribute: 'show-max',
    type: Number
  })
  public showMax = 6;

  /**
   * Sets whether the person image should be fetched
   * from the Microsoft Graph
   *
   * @type {boolean}
   * @memberof MgtPerson
   */
  @property({
    attribute: 'disable-images',
    type: Boolean
  })
  public disableImages = false;

  @property({
    attribute: 'show-presence',
    type: Boolean
  })
  public showPresence: boolean;

  /**
   * Sets how the person-card is invoked
   * Set to PersonCardInteraction.none to not show the card
   *
   * @type {PersonCardInteraction}
   * @memberof MgtPerson
   */
  @property({
    attribute: 'person-card',
    converter: value => personCardConverter(value)
  })
  public personCardInteraction: PersonCardInteraction = 'none';

  /**
   * array of user picked people.
   *
   * @type {IDynamicPerson[]}
   */
  public get selectedPeople(): IDynamicPerson[] {
    return this._selectedPeople;
  }
  @property({
    attribute: 'selected-people',
    type: Array
  })
  public set selectedPeople(value: IDynamicPerson[]) {
    if (!value) value = [];
    if (!arraysAreEqual(this._selectedPeople, value)) {
      this._selectedPeople = value;
      this.fireCustomEvent('selectionChanged', this._selectedPeople);
    }
  }

  /**
   * array of people to be selected upon initialization
   *
   * @type {string[]}
   * @memberof MgtPeoplePicker
   */
  @property({
    attribute: 'default-selected-user-ids',
    converter: value => {
      return value.split(',').map(v => v.trim());
    },
    type: String
  })
  public defaultSelectedUserIds: string[] = [];

  /**
   * array of groups to be selected upon initialization
   *
   * @type {string[]}
   * @memberof MgtPeoplePicker
   */
  @property({
    attribute: 'default-selected-group-ids',
    converter: value => {
      return value.split(',').map(v => v.trim());
    },
    type: String
  })
  public defaultSelectedGroupIds: string[] = [];

  /**
   * Placeholder text.
   *
   * @type {string}
   * @memberof MgtPeoplePicker
   */
  @property({
    attribute: 'placeholder',
    type: String
  })
  public placeholder: string;

  /**
   * Determines whether component should be disabled or not
   *
   * @type {boolean}
   * @memberof MgtPeoplePicker
   */
  @property({
    attribute: 'disabled',
    type: Boolean
  })
  public disabled = false;

  /**
   * Determines if a user can enter an email without selecting a person
   *
   * @type {boolean}
   * @memberof MgtPeoplePicker
   */
  @property({
    attribute: 'allow-any-email',
    type: Boolean
  })
  public allowAnyEmail = false;

  /**
   * Determines whether component allows multiple or single selection of people
   *
   * @type {string}
   * @memberof MgtPeoplePicker
   */
  @property({
    attribute: 'selection-mode',
    type: String
  })
  public selectionMode: string;

  /**
   * Array of the only users to be searched.
   *
   * @type {string[]}
   * @memberof MgtPeoplePicker
   */
  @property({
    attribute: 'user-ids',
    converter: value => {
      return value.split(',').map(v => v.trim());
    },
    type: String
  })
  public userIds: string[] = [];

  /**
   * Filters that can be set on the user properties query.
   */
  @property({ attribute: 'user-filters' })
  public userFilters: string;

  /**
   * Filters that can be set on the people query properties.
   */
  @property({ attribute: 'people-filters' })
  public peopleFilters: string;

  /**
   * Filters that can be set on the group query properties.
   */
  @property({ attribute: 'group-filters' })
  public groupFilters: string;

  /**
   * Label that can be set on the people picker input to provide context to
   * assistive technologies
   */
  @property({
    attribute: 'aria-label',
    type: String
  })
  public ariaLabel: string;

  /**
   * Sets whether the people suggestions should apper on the suggestion list
   *
   * @type {boolean}
   * @memberof MgtPerson
   */
  @property({
    attribute: 'disable-suggestions',
    type: Boolean
  })
  public disableSuggestions = false;

  /**
   * Get the scopes required for people picker
   *
   * @static
   * @return {*}  {string[]}
   * @memberof MgtPeoplePicker
   */
  public static get requiredScopes(): string[] {
    return [
      ...new Set([
        'user.read.all',
        'people.read',
        'group.read.all',
        'user.readbasic.all',
        'contacts.read',
        ...MgtPerson.requiredScopes
      ])
    ];
  }

  /**
   * User input in search.
   *
   * @protected
   * @type {string}
   * @memberof MgtPeoplePicker
   */
  @state()
  protected userInput = '';

  private _selectedPeople: IDynamicPerson[] = [];

  private defaultPeople: IDynamicPerson[];

  // tracking of user arrow key input for selection
  @state() private _arrowSelectionCount = -1;
  // List of people requested if group property is provided
  private _groupPeople: IDynamicPerson[];
  private _debouncedSearch: { (): void; (): void };
  private defaultSelectedUsers: IDynamicPerson[] = [];
  private defaultSelectedGroups: IDynamicPerson[] = [];
  // List of users highlighted for copy/cut-pasting
  @state() private _highlightedUsers: Element[] = [];
  // current user index to the left of the highlighted users
  private _currentHighlightedUserPos = 0;

  /**
   * Checks if the input is focused.
   */
  @state() private _isFocused = false;

  /**
   * Switch to determine if a typed email can be set.
   */
  @state() private _setAnyEmail = false;

  /**
   * List of people found from the graph calls.
   */
  @state() private _foundPeople: IDynamicPerson[];

  constructor() {
    super();

    this.addEventListener('copy', this.handleCopy);
    this.addEventListener('cut', this.handleCut);
    this.addEventListener('paste', this.handlePaste);
    this.addEventListener('selectionChanged', this.handleSelectionChanged);
  }

  /**
   * Disable the inner input of the fluent-text-field.
   */
  private disableTextInput() {
    const inputControl = this.input.shadowRoot.querySelector<HTMLInputElement>('input');
    if (inputControl) {
      inputControl.setAttribute('disabled', 'true');
      inputControl.value = '';
    }
  }

  /**
   * Enable the inner input of the fluent-text-field.
   */
  private enableTextInput() {
    const inputControl = this.input.shadowRoot.querySelector<HTMLInputElement>('input');
    if (inputControl) {
      inputControl.removeAttribute('disabled');
      inputControl.focus();
    }
  }

  /**
   * Clears the disabled property on the people picker when used in single mode.
   */
  private readonly handleSelectionChanged = () => {
    if (this.selectedPeople.length === 0 && !this.disabled) {
      this.enableTextInput();
    }
  };

  private get hasMaxSelections(): boolean {
    return this.selectionMode === 'single' && this.selectedPeople.length >= 1;
  }

  /**
   * Focuses the input element when focus is called
   *
   * @param {FocusOptions} [options]
   * @memberof MgtPeoplePicker
   */
  public focus(options?: FocusOptions) {
    if (!this.input) {
      return;
    }
    this.input.focus(options);
    this.input.select();
  }

  /**
   * Queries the microsoft graph for a user based on the user id and adds them to the selectedPeople array
   *
   * @param {readonly string []} an array of user ids to add to selectedPeople
   * @returns {Promise<void>}
   * @memberof MgtPeoplePicker
   */
  public async selectUsersById(userIds: readonly string[]): Promise<void> {
    const provider = Providers.globalProvider;
    const graph = Providers.globalProvider.graph;
    if (provider && provider.state === ProviderState.SignedIn) {
      // eslint-disable-next-line guard-for-in, @typescript-eslint/no-for-in-array
      for (const id in userIds) {
        const userId = userIds[id];
        try {
          const personDetails = await getUser(graph, userId, defaultPersonProperties);
          this.addPerson(personDetails);
        } catch (e: unknown) {
          // This caters for allow-any-email property if it's enabled on the component
          // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-unsafe-call
          if (isGraphError(e) && e.message?.includes('does not exist') && this.allowAnyEmail) {
            if (isValidEmail(userId)) {
              const anyMailUser = {
                mail: userId,
                displayName: userId
              };
              this.addPerson(anyMailUser);
            }
          }
        }
      }
    }
  }
  /**
   * Queries the microsoft graph for a group of users from a group id, and adds them to the selectedPeople
   *
   * @param {readonly string []} an array of group ids to add to selectedPeople
   * @returns {Promise<void>}
   * @memberof MgtPeoplePicker
   */
  public async selectGroupsById(groupIds: readonly string[]): Promise<void> {
    const provider = Providers.globalProvider;
    const graph = Providers.globalProvider.graph;
    if (provider && provider.state === ProviderState.SignedIn) {
      // eslint-disable-next-line guard-for-in, @typescript-eslint/no-for-in-array
      for (const id in groupIds) {
        try {
          const groupDetails = await getGroup(graph, groupIds[id]);
          this.addPerson(groupDetails);
        } catch (e) {
          // no-op
        }
      }
    }
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return a lit-html TemplateResult.
   * Setting properties inside this method will not trigger the element to update.
   *
   * @returns {TemplateResult}
   * @memberof MgtPeoplePicker
   */
  public render(): TemplateResult {
    const defaultTemplate = this.renderTemplate('default', { people: this._foundPeople });
    if (defaultTemplate) {
      return defaultTemplate;
    }

    const selectedPeopleTemplate = this.renderSelectedPeople(this.selectedPeople);

    if (this.selectionMode === 'single' && this.selectedPeople.length === 1) {
      return selectedPeopleTemplate;
    } else {
      const inputTemplate = this.renderInput(selectedPeopleTemplate);
      const flyoutTemplate = this.renderFlyout(inputTemplate);

      return html`
      <div>
        ${flyoutTemplate}
      </div>
    `;
    }
  }

  protected args(): unknown[] {
    return [
      this.providerState,
      this.groupId,
      this.groupIds,
      this.type,
      this.groupType,
      this.userType,
      this.transitiveSearch,
      this.people,
      this.showMax,
      this.defaultSelectedUserIds,
      this.defaultSelectedGroupIds,
      this.userIds,
      this.userFilters,
      this.peopleFilters,
      this.groupFilters
    ];
  }

  /**
   * Clears state of the component
   *
   * @protected
   * @memberof MgtPeoplePicker
   */
  protected clearState(): void {
    this.selectedPeople = [];
    this.userInput = '';
    this._highlightedUsers = [];
    this._currentHighlightedUserPos = 0;
  }

  /**
   * Render the input text box.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPeoplePicker
   */
  protected renderInput(selectedPeopleTemplate: TemplateResult): TemplateResult {
    const placeholder = this.disabled ? '' : this.placeholder || this.strings.inputPlaceholderText;
    const maxSelectionsAriaLabel = this.hasMaxSelections ? this.strings.maxSelectionsAriaLabel : '';

    const searchIcon = html`<span class="search-icon">${getSvg(SvgIcon.Search)}</span>`;
    const startSlot = this.selectedPeople?.length > 0 ? selectedPeopleTemplate : searchIcon;
    return html`
      <fluent-text-field
        autocomplete="off"
        appearance="outline"
        slot="anchor"
        id="people-picker-input"
        role="combobox"
        aria-live="assertive"
        aria-atomic="true"
        placeholder=${this.hasMaxSelections ? this.strings.maxSelectionsPlaceHolder : placeholder}
        aria-label=${this.ariaLabel || maxSelectionsAriaLabel || placeholder || this.strings.selectContact}
        aria-expanded=${this.flyout?.isOpen ?? false}
        @click="${this.hasMaxSelections ? undefined : this.handleInputClick}"
        @focus="${this.hasMaxSelections ? undefined : this.gainedFocus}"
        @keydown="${this.hasMaxSelections ? undefined : this.onUserKeyDown}"
        @input="${this.hasMaxSelections ? undefined : this.onUserInput}"
        @blur="${this.lostFocus}"
        ?disabled=${this.disabled}
      >
        <span slot="start">${startSlot}</span>
      </fluent-text-field>
    `;
  }

  /**
   * Render the selected people tokens.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPeoplePicker
   */
  protected renderSelectedPeople(selectedPeople?: IDynamicPerson[]): TemplateResult {
    if (!selectedPeople?.length) {
      return html``;
    }

    return html`
       <ul
        id="selected-list"
        aria-label="${this.strings.selected}"
        class="selected-list">
          ${repeat(
            selectedPeople,
            person => person?.id,
            person => html`
            <li class="selected-list-item">
              ${
                this.renderTemplate(
                  'selected-person',
                  { person },
                  `selected-${person?.id ? person.id : person.displayName}`
                ) || this.renderSelectedPerson(person)
              }

              <div
                role="button"
                tabindex="0"
                class="selected-list-item-close-icon"
                aria-label="${this.strings.removeSelectedUser}${person?.displayName ?? ''}"
                @click="${(e: UIEvent) => this.removePerson(person, e)}"
                @keydown="${(e: KeyboardEvent) => this.handleRemovePersonKeyDown(person, e)}">
                  ${getSvg(SvgIcon.Close)}
              </div>
          </li>`
          )}
      </ul>`;
  }
  /**
   * Render the flyout chrome.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPeoplePicker
   */
  protected renderFlyout(anchor: TemplateResult): TemplateResult {
    return mgtHtml`
       <mgt-flyout light-dismiss class="flyout">
         ${anchor}
         <fluent-card
          tabindex="0"
          slot="flyout"
          class="flyout-root"
          @wheel=${(e: WheelEvent) => this.handleSectionScroll(e)}
          @keydown=${(e: KeyboardEvent) => this.onUserKeyDown(e)}
          class="custom">
           ${this.renderFlyoutContent()}
         </fluent-card>
       </mgt-flyout>
     `;
  }

  /**
   * Render the appropriate state in the results flyout.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPeoplePicker
   */
  protected renderFlyoutContent(): TemplateResult {
    if (this._task.status === TaskStatus.PENDING) {
      return this.renderLoading();
    }

    const people = this._foundPeople;

    if (!people || people.length === 0 || this.showMax === 0) {
      return this.renderNoData();
    } else {
      return this.renderSearchResults(people);
    }
  }

  /**
   * Render the loading state.
   *
   * @protected
   * @returns
   * @memberof MgtPeoplePicker
   */
  protected renderLoading = (): TemplateResult => {
    return (
      this.renderTemplate('loading', null) ||
      mgtHtml`
         <div class="message-parent">
           <mgt-spinner></mgt-spinner>
           <div aria-label="${this.strings.loadingMessage}" aria-live="polite" role="alert" class="loading-text">
             ${this.strings.loadingMessage}
           </div>
         </div>
       `
    );
  };

  /**
   * Render the state when no results are found for the search query.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPeoplePicker
   */
  protected renderNoData(): TemplateResult {
    if (!this._isFocused) {
      return;
    }
    return (
      this.renderTemplate('error', null) ||
      this.renderTemplate('no-data', null) ||
      html`
         <div class="message-parent">
           <div aria-label=${this.strings.noResultsFound} aria-live="polite" role="alert" class="search-error-text">
             ${this.strings.noResultsFound}
           </div>
         </div>
       `
    );
  }

  /**
   * Render the list of search results.
   *
   * @protected
   * @param {IDynamicPerson[]} people
   * @returns
   * @memberof MgtPeoplePicker
   */
  protected renderSearchResults(people: IDynamicPerson[]) {
    const filteredPeople = people.filter(person => person.id);
    return html`
      <ul
        id="suggestions-list"
        class="searched-people-list"
        role="listbox"
        aria-live="polite"
        aria-label="${this.strings.suggestionsTitle}"
        title=${this.strings.suggestionsTitle}
      >
        ${repeat(
          filteredPeople,
          person => person.id,
          person => html`
          <li
            id="${person.id}"
            class="searched-people-list-result"
            role="option"
            @click="${() => this.handleSuggestionClick(person)}">
              ${this.renderPersonResult(person)}
          </li>
        `
        )}
      </ul>
     `;
  }

  /**
   * Render an individual person search result.
   *
   * @protected
   * @param {IDynamicPerson} person
   * @returns {TemplateResult}
   * @memberof MgtPeoplePicker
   */
  protected renderPersonResult(person: IDynamicPerson): TemplateResult {
    return (
      this.renderTemplate('person', { person }, person.id) ||
      mgtHtml`
         <mgt-person
          class="person-image-result"
          ?show-presence=${this.showPresence}
          view="twolines"
          line2-property="jobTitle,mail"
          .personDetails=${person}
          ?fetch-image=${!this.disableImages}
          person-card="none"
        ></mgt-person>`
    );
  }

  /**
   * Render an individual selected person token.
   *
   * @protected
   * @param {IDynamicPerson} person
   * @returns {TemplateResult}
   * @memberof MgtPeoplePicker
   */
  protected renderSelectedPerson(person: IDynamicPerson): TemplateResult {
    return mgtHtml`
       <mgt-person
         class="person-image-selected"
         .personDetails=${person}
         ?fetch-image=${!this.disableImages}
         view="oneline"
         person-card=${this.personCardInteraction}
        ></mgt-person>
     `;
  }

  /**
   * Async query to Graph for members of group if determined by developer.
   * set's `this.groupPeople` to those members.
   */
  protected async loadState(): Promise<void> {
    this._groupPeople = null;
    this._foundPeople = null;
    this.defaultPeople = null;
    let people = this.people;
    const input = this.userInput.toLowerCase();
    const provider = Providers.globalProvider;

    if (people?.length) {
      if (input) {
        const displayNameMatch = people.filter(person => person?.displayName.toLowerCase().includes(input));
        people = displayNameMatch;
      }
    } else if (!people && provider && provider.state === ProviderState.SignedIn) {
      const graph = provider.graph.forComponent(this);

      if (!input.length) {
        if (
          (this.defaultSelectedUserIds.length > 0 || this.defaultSelectedGroupIds.length > 0) &&
          !this.selectedPeople.length &&
          !this.defaultSelectedUsers.length &&
          !this.defaultSelectedGroups.length
        ) {
          this.defaultSelectedUsers = await getUsersForUserIds(
            graph,
            this.defaultSelectedUserIds,
            '',
            this.userFilters
          );
          this.defaultSelectedGroups = await getGroupsForGroupIds(
            graph,
            this.defaultSelectedGroupIds,
            this.peopleFilters
          );

          this.defaultSelectedGroups = this.defaultSelectedGroups.filter(group => {
            return group !== null;
          });

          this.defaultSelectedUsers = this.defaultSelectedUsers.filter(user => {
            return user !== null;
          });

          this.selectedPeople = [...this.defaultSelectedUsers, ...this.defaultSelectedGroups];

          if (this.hasMaxSelections) {
            this.disableTextInput();
          }
          this.requestUpdate();
        }
        if (this.disableSuggestions) {
          this._foundPeople = [];
          return;
        }
        if (this.defaultPeople) {
          people = this.defaultPeople;
        } else {
          if (this.groupId || this.groupIds.length) {
            if (this._groupPeople === null) {
              if (this.groupId) {
                try {
                  if (this.type === 'group') {
                    this._groupPeople = await findGroupMembers(
                      graph,
                      null,
                      this.groupId,
                      this.showMax,
                      this.type,
                      this.transitiveSearch
                    );
                  } else {
                    this._groupPeople = await findGroupMembers(
                      graph,
                      null,
                      this.groupId,
                      this.showMax,
                      this.type,
                      this.transitiveSearch,
                      this.userFilters,
                      this.peopleFilters
                    );
                  }
                } catch (_) {
                  this._groupPeople = [];
                }
              } else {
                if (this.type === 'group') {
                  try {
                    this._groupPeople = await getGroupsForGroupIds(graph, this.groupIds, this.groupFilters);
                  } catch (_) {
                    this._groupPeople = [];
                  }
                } else {
                  try {
                    const peopleInGroups = await findUsersFromGroupIds(
                      graph,
                      '',
                      this.groupIds,
                      this.showMax,
                      this.type,
                      this.transitiveSearch,
                      this.userFilters
                    );
                    this._groupPeople = peopleInGroups;
                  } catch (_) {
                    this._groupPeople = [];
                  }
                }
              }
            }
            people = this._groupPeople || [];
          } else if (this.type === 'person' || this.type === 'any') {
            if (this.userIds.length) {
              people = await getUsersForUserIds(graph, this.userIds, '', this.userFilters);
            } else {
              const isUserOrContactType = this.userType === 'user' || this.userType === 'contact';
              if (this.userFilters && isUserOrContactType) {
                people = await getUsers(graph, this.userFilters, this.showMax);
              } else {
                people = await getPeople(graph, this.userType, this.peopleFilters, this.showMax);
              }
            }
          } else if (this.type === 'group') {
            let groups = (await findGroups(graph, '', this.showMax, this.groupType, this.groupFilters)) || [];
            // eslint-disable-next-line @typescript-eslint/dot-notation
            if (groups.length > 0 && groups[0]['value']) {
              // eslint-disable-next-line @typescript-eslint/dot-notation, @typescript-eslint/no-unsafe-assignment
              groups = groups[0]['value'];
            }
            people = groups;
          }
          this.defaultPeople = people;
        }
      }

      if (input) {
        people = [];

        if (this.groupId) {
          people =
            (await findGroupMembers(
              graph,
              input,
              this.groupId,
              this.showMax,
              this.type,
              this.transitiveSearch,
              this.userFilters,
              this.peopleFilters
            )) || [];
        } else {
          if (this.type === 'person' || this.type === 'any') {
            try {
              // Default UserType === any
              if (this.userType === 'contact' || this.userType === 'user') {
                // we might have a user-filters property set, search for users with it.
                if (this.userIds?.length) {
                  // has the user-ids property set
                  people = await getUsersForUserIds(graph, this.userIds, input, this.userFilters);
                } else {
                  people = await findUsers(graph, input, this.showMax, this.userFilters);
                }
              } else {
                if (!this.groupIds.length) {
                  if (this.userIds?.length) {
                    // has the user-ids property set
                    people = await getUsersForUserIds(graph, this.userIds, input, this.userFilters);
                  } else {
                    people = (await findPeople(graph, input, this.showMax, this.userType, this.peopleFilters)) || [];
                  }
                } else {
                  // Does not work when the PersonType = person.
                  try {
                    people = await findUsersFromGroupIds(
                      graph,
                      input,
                      this.groupIds,
                      this.showMax,
                      this.type,
                      this.transitiveSearch,
                      this.userFilters
                    );
                  } catch (_) {
                    // nop
                  }
                }
              }
            } catch (e) {
              // nop
            }

            // Don't follow this path if a people-filters attribute is set on the component as the
            // default type === PersonType.person
            if (people && people.length < this.showMax && this.userType !== 'contact' && this.type !== 'person') {
              try {
                const users = (await findUsers(graph, input, this.showMax, this.userFilters)) || [];

                // make sure only unique people
                const peopleIds = new Set(people.map(p => p.id));
                for (const user of users) {
                  if (!peopleIds.has(user.id)) {
                    people.push(user);
                  }
                }
              } catch (e) {
                // no-op
              }
            }
          }

          if ((this.type === 'group' || this.type === 'any') && people.length < this.showMax) {
            if (this.groupIds.length) {
              try {
                people = await findGroupsFromGroupIds(
                  graph,
                  input,
                  this.groupIds,
                  this.showMax,
                  this.groupType,
                  this.userFilters
                );
              } catch (_) {
                // no-op
              }
            } else {
              let groups = [];
              try {
                groups = (await findGroups(graph, input, this.showMax, this.groupType, this.groupFilters)) || [];
                people = people.concat(groups);
              } catch (e) {
                // nop
              }
            }
          }
        }
      }
    }
    // people = this.getUniquePeople(people);
    this._foundPeople = this.filterPeople(people);
  }

  /**
   * Gets the Groups in a list of group IDs.
   *
   * @param graph the graph object
   * @param people already found groups
   * @returns groups found
   */
  private async getGroupsForGroupIds(graph: IGraph, people: IDynamicPerson[]) {
    const groups = await getGroupsForGroupIds(graph, this.groupIds, this.groupFilters);
    for (const group of groups as IDynamicPerson[]) {
      people = people.concat(group);
    }
    people = people.filter(person => person);
    return people;
  }

  /**
   * Hide the results flyout.
   *
   * @protected
   * @memberof MgtPeoplePicker
   */
  protected hideFlyout(): void {
    const flyout = this.flyout;
    if (flyout) {
      flyout.close();
    }
    if (this.input) {
      this.input.setAttribute('aria-activedescendant', '');
    }
    this._arrowSelectionCount = -1;
  }

  /**
   * Show the results flyout.
   *
   * @protected
   * @memberof MgtPeoplePicker
   */
  protected showFlyout(): void {
    const flyout = this.flyout;
    if (flyout) {
      flyout.open();
    }
    this._arrowSelectionCount = -1;
  }

  /**
   * Removes person from selected people
   *
   * @param person - person and details pertaining to user selected
   */
  protected removePerson(person: IDynamicPerson, e: UIEvent): void {
    e.stopPropagation();
    const filteredPersonArr = this.selectedPeople.filter(p => {
      if (!person.id && p.displayName) {
        return p.displayName !== person.displayName;
      }
      return p.id !== person.id;
    });
    this.selectedPeople = filteredPersonArr;
    void this.loadState();
  }

  /**
   * Checks if key pressed is an `Enter` key before removing person
   *
   * @param person
   * @param e
   */
  protected handleRemovePersonKeyDown(person: IDynamicPerson, e: KeyboardEvent): void {
    if (e.key === 'Enter') {
      this.removePerson(person, e);
      this.enableTextInput();
    }
  }

  /**
   * Tracks when user selects person from picker
   *
   * @param person - contains details pertaining to selected user
   * @param event - tracks user event
   */
  protected addPerson(person: IDynamicPerson): void {
    if (person) {
      setTimeout(() => {
        this.clearInput();
      }, 50);
      const duplicatePeople = this.selectedPeople.filter(p => {
        if (!person.id && p.displayName) {
          return p.displayName === person.displayName;
        }
        return p.id === person.id;
      });

      if (duplicatePeople.length === 0) {
        this.selectedPeople = [...this.selectedPeople, person];
        void this.loadState();
        this._foundPeople = [];
        this._arrowSelectionCount = -1;
      }
      if (this._isFocused) {
        this.enableTextInput();
      }
    }
  }

  private clearInput() {
    this.clearHighlighted();
    if (this.selectionMode !== 'single') {
      this.input.value = '';
    }
    this.userInput = '';
  }

  // handle input click
  private readonly handleInputClick = () => {
    if (!this.flyout.isOpen) {
      this.handleUserSearch();
    }
  };

  // handle input focus
  private readonly gainedFocus = () => {
    this.clearHighlighted();
    this._isFocused = true;
    void this.loadState();
    this.showFlyout();
  };

  // handle input blur
  private readonly lostFocus = () => {
    this._isFocused = false;
    if (this.input) {
      this.input.setAttribute('aria-activedescendant', '');
    }

    const peopleList = this.renderRoot.querySelector('.people-list');

    if (peopleList) {
      for (const el of peopleList.children) {
        el.classList.remove('focused');
        el.setAttribute('aria-selected', 'false');
      }
    }

    this.requestUpdate();
  };

  /**
   * Handles input from the key up events on the keyboard.
   */
  private readonly onUserKeyUp = (event: KeyboardEvent): void => {
    const keyName = event.key;
    const isCmdOrCtrlKey = event.getModifierState('Control') || event.getModifierState('Meta');
    const isPaste = isCmdOrCtrlKey && keyName === 'v';
    const isArrowKey = ['ArrowDown', 'ArrowRight', 'ArrowUp', 'ArrowLeft'].includes(keyName);

    if ((!isPaste && isCmdOrCtrlKey) || isArrowKey) {
      if (isCmdOrCtrlKey || ['ArrowLeft', 'ArrowRight'].includes(keyName)) {
        // Only hide the flyout when you're doing selections with Left/Right Arrow key
        this.hideFlyout();
      }

      if (keyName === 'ArrowDown') {
        if (!this.flyout.isOpen && this._isFocused) {
          this.handleUserSearch();
        }
      }
      return;
    }

    if (['Tab', 'Enter', 'Shift'].includes(keyName)) return;

    if (keyName === 'Escape') {
      this.clearInput();
      this._foundPeople = [];
      this._arrowSelectionCount = -1;
      return;
    }

    if (keyName === 'Backspace' && this.userInput.length === 0 && this.selectedPeople.length > 0) {
      this.clearHighlighted();
      // remove last person in selected list
      this.selectedPeople = this.selectedPeople.splice(0, this.selectedPeople.length - 1);
      void this.loadState();
      this.hideFlyout();
      return;
    }

    if ([';', ','].includes(keyName)) {
      if (this.allowAnyEmail) {
        this._setAnyEmail = true;
        event.preventDefault();
        event.stopPropagation();
      }
      return;
    }
  };

  private readonly onUserInput = (event: InputEvent) => {
    const input = event.target as HTMLInputElement;
    this.userInput = input.value;
    if (this.userInput) {
      const validEmail = isValidEmail(this.userInput);
      if (validEmail && this.allowAnyEmail) {
        if (this._setAnyEmail) {
          this.handleAnyEmail();
        }
      } else {
        this.handleUserSearch();
      }
      this._setAnyEmail = false;
    } else {
      this._foundPeople = [];
    }
  };

  private handleAnyEmail() {
    this._arrowSelectionCount = -1;
    if (isValidEmail(this.userInput)) {
      const anyMailUser = {
        mail: this.userInput,
        displayName: this.userInput
      };
      this.addPerson(anyMailUser);
    }
    this.hideFlyout();
    if (this.input) {
      this.input.focus();
      this._isFocused = true;
    }
  }

  // handle suggestion list item click
  private handleSuggestionClick(person: IDynamicPerson): void {
    this.addPerson(person);
    if (this.hasMaxSelections) {
      this.disableTextInput();
      this.input.value = '';
    }
    this.hideFlyout();
  }

  /**
   * Tracks event on user input in search
   *
   * @param input - input text
   */
  private handleUserSearch() {
    if (!this._debouncedSearch) {
      this._debouncedSearch = debounce(async () => {
        await this._task.run();
        this._arrowSelectionCount = -1;
        this.showFlyout();
      }, 400);
    }

    this._debouncedSearch();
  }

  /**
   * Tracks event on user search (keydown)
   *
   * @param event - event tracked on user input (keydown)
   */
  private readonly onUserKeyDown = (event: KeyboardEvent): void => {
    event.stopPropagation();
    const keyName = event.key;
    const selectedList = this.renderRoot.querySelector('.selected-list');
    const isCmdOrCtrlKey = event.getModifierState('Control') || event.getModifierState('Meta');
    if (isCmdOrCtrlKey && selectedList) {
      const selectedPeople = selectedList.querySelectorAll('mgt-person.person-image-selected');
      this.hideFlyout();
      if (isCmdOrCtrlKey && keyName === 'ArrowLeft') {
        this._currentHighlightedUserPos =
          (this._currentHighlightedUserPos - 1 + selectedPeople.length) % selectedPeople.length;
        if (this._currentHighlightedUserPos >= 0 && !Number.isNaN(this._currentHighlightedUserPos)) {
          this._highlightedUsers.push(selectedPeople[this._currentHighlightedUserPos]);
        } else {
          this._currentHighlightedUserPos = 0;
        }
      } else if (isCmdOrCtrlKey && keyName === 'ArrowRight') {
        const person = this._highlightedUsers.pop();
        if (person) {
          const personParent = person.parentElement;
          if (personParent) {
            this.clearHighlighted(personParent);
            this._currentHighlightedUserPos++;
          }
        }
      } else if (isCmdOrCtrlKey && keyName === 'a') {
        this._highlightedUsers = [];
        selectedPeople.forEach(person => this._highlightedUsers.push(person));
      }
      if (this._highlightedUsers) {
        this.highlightSelectedPeople(this._highlightedUsers);
      }
      return;
    }

    this.clearHighlighted();

    if (!this.flyout.isOpen) {
      return;
    }

    if (keyName === 'ArrowUp' || keyName === 'ArrowDown') {
      this.handleArrowSelection(event);
      event.preventDefault();
    }

    if (keyName === 'Enter') {
      if (!event.shiftKey && this._foundPeople) {
        event.preventDefault();

        const foundPerson = this._foundPeople[this._arrowSelectionCount];
        if (foundPerson) {
          this.addPerson(foundPerson);
          this.hideFlyout();
          this.input.value = '';
          if (this.hasMaxSelections) {
            this.disableTextInput();
          }
          return;
        }
      }

      if (this.allowAnyEmail) {
        this.handleAnyEmail();
      } else {
        this.showFlyout();
      }
    }

    if ([';', ',', 'Tab'].includes(keyName)) {
      if (this.allowAnyEmail) {
        // need to ensure the tab key does tab things
        if ('Tab' !== keyName) event.preventDefault();
        this.userInput = this.input.value;
        this.handleAnyEmail();
      }
    }
    // need to ensure that the tab key hides the flyout even if the input doesn't have an email address
    if (keyName === 'Tab') {
      this.hideFlyout();
    }
  };

  /**
   * Gets the text of the highlighed people and writes it to the clipboard
   */
  private async writeHighlightedText() {
    const copyText = [];
    for (const element of this._highlightedUsers) {
      // eslint-disable-next-line @typescript-eslint/dot-notation
      const { id, displayName, mail, userPrincipalName, scoredEmailAddresses } = element['_personDetails'] as Person &
        User;
      let emailAddress: string;
      if (scoredEmailAddresses && scoredEmailAddresses.length > 0) {
        emailAddress = scoredEmailAddresses.pop().address;
      } else {
        emailAddress = userPrincipalName || mail;
      }

      copyText.push({ id, displayName, email: emailAddress });
    }
    let copiedTextStr = '';
    if (copyText.length > 0) {
      copiedTextStr = JSON.stringify(copyText);
    }

    await navigator.clipboard.writeText(copiedTextStr);
  }

  /**
   * Handles the cut event when it is fired
   */
  private readonly handleCut = () => {
    this.writeHighlightedText().then(
      () => {
        this.removeHighlightedOnCut();
      },
      () => {
        // intentionally left blank
      }
    );
  };

  /**
   * Handles the copy event when it is fired
   */
  private readonly handleCopy = () => {
    void this.writeHighlightedText();
  };

  /**
   * Parses the copied people text and adds them when you paste
   */
  private readonly handlePaste = () => {
    navigator.clipboard.readText().then(
      copiedText => {
        if (copiedText) {
          try {
            const people: IDynamicPerson[] = JSON.parse(copiedText) as IDynamicPerson[];
            if (people && people.length > 0) {
              for (const person of people) {
                this.addPerson(person);
              }
            }
          } catch (error) {
            if (error instanceof SyntaxError) {
              const delimiters = [',', ';'];
              let listOfUsers: string[];
              try {
                for (const delimiter of delimiters) {
                  listOfUsers = copiedText.split(delimiter);
                  if (listOfUsers.length > 1) {
                    this.hideFlyout();
                    void this.selectUsersById(listOfUsers);
                    break;
                  }
                }
                // eslint-disable-next-line no-empty
              } catch (_) {}
            }
          }
        }
      },
      // eslint-disable-next-line @typescript-eslint/no-unused-vars
      error => {
        // 'navigator.clipboard.readText is not a function' error is thrown in Mozilla
        // more information here https://developer.mozilla.org/en-US/docs/Web/API/Clipboard/readText#browser_compatibility
        // Firefox only supports reading the clipboard in browser extensions,
        // using the "clipboardRead" extension permission.
      }
    );
  };

  /**
   * Removes only the highlighted elements from the peoplePicker during cut operations.
   */
  private removeHighlightedOnCut() {
    this.selectedPeople = this.selectedPeople.splice(0, this.selectedPeople.length - this._highlightedUsers.length);
    this._highlightedUsers = [];
    this._currentHighlightedUserPos = 0;
    void this.loadState();
    this.hideFlyout();
  }
  /**
   * Changes the color class to show which people are selected for copy/cut-paste
   *
   * @param people list of selected people classes
   */
  private highlightSelectedPeople(people: Element[]) {
    for (const person of people) {
      const parentElement = person?.parentElement;
      parentElement.classList.add('highlighted');
    }
  }

  /**
   * Defaults the people class back to the normal view
   */
  private clearHighlighted(node?: Element) {
    if (node) {
      node.classList.remove('highlighted');
    } else {
      for (const person of this._highlightedUsers) {
        const parentElement = person.parentElement;
        if (parentElement) {
          parentElement.classList.remove('highlighted');
        }
      }
      this._highlightedUsers = [];
      this._currentHighlightedUserPos = 0;
    }
  }

  /**
   * Tracks user key selection for arrow key selection of people
   *
   * @param event - tracks user key selection
   */
  private handleArrowSelection(event?: KeyboardEvent): void {
    const peopleList = this.renderRoot.querySelector('.searched-people-list');

    if (peopleList?.children?.length) {
      if (event) {
        // update arrow count
        if (event.key === 'ArrowUp') {
          if (this._arrowSelectionCount === -1) {
            this._arrowSelectionCount = 0;
          } else {
            this._arrowSelectionCount =
              (this._arrowSelectionCount - 1 + peopleList.children.length) % peopleList.children.length;
          }
        }
        if (event.key === 'ArrowDown') {
          if (this._arrowSelectionCount === -1) {
            this._arrowSelectionCount = 0;
          } else {
            this._arrowSelectionCount =
              (this._arrowSelectionCount + 1 + peopleList.children.length) % peopleList.children.length;
          }
        }
      }

      for (const person of peopleList?.children ?? []) {
        const p = person as HTMLElement;
        p.setAttribute('aria-selected', 'false');
        p.blur();
        p.removeAttribute('tabindex');
      }

      // set selected background
      // set aria-selected to true
      const focusedItem = peopleList.children[this._arrowSelectionCount] as HTMLElement;

      if (focusedItem) {
        focusedItem.setAttribute('tabindex', '0');
        focusedItem.focus();
        focusedItem.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'nearest' });
        focusedItem.setAttribute('aria-selected', 'true');
        this.input.setAttribute('aria-activedescendant', focusedItem?.id);
      }
    }
  }

  /**
   * Filters people searched from already selected people
   *
   * @param people - array of people returned from query to Graph
   */
  private filterPeople(people: IDynamicPerson[]): IDynamicPerson[] {
    // check if people need to be updated
    // ensuring people list is displayed
    // find ids from selected people
    const uniquePeople: IDynamicPerson[] = [];
    if (people && people.length > 0) {
      people = people.filter(person => person);
      const idFilter = this.selectedPeople.map(el => {
        return el.id ? el.id : el.displayName;
      });

      // filter id's
      const filtered = people.filter((person: IDynamicPerson) => {
        if (person?.id) {
          return idFilter.indexOf(person.id) === -1;
        } else {
          return idFilter.indexOf(person?.displayName) === -1;
        }
      });

      // remove duplicates
      const dupsSet = new Set<string>();
      for (const d of filtered) {
        const person = JSON.stringify(d);
        dupsSet.add(person);
      }

      dupsSet.forEach((person: string) => {
        const p: IDynamicPerson = JSON.parse(person) as IDynamicPerson;
        uniquePeople.push(p);
      });
    }
    return uniquePeople;
  }

  // stop propagating wheel event to flyout so mouse scrolling works
  private handleSectionScroll(e: WheelEvent) {
    const target = this.renderRoot.querySelector('.flyout-root');
    if (target) {
      if (
        !(e.deltaY < 0 && target.scrollTop === 0) &&
        !(e.deltaY > 0 && target.clientHeight + target.scrollTop >= target.scrollHeight - 1)
      ) {
        e.stopPropagation();
      }
    }
  }
}
