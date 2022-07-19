/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, internalProperty, property, TemplateResult } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { MgtTemplatedComponent, Providers, ProviderState, TeamsHelper } from '@microsoft/mgt-element';
import { Presence, User, Person } from '@microsoft/microsoft-graph-types';

import { findPeople, getEmailFromGraphEntity } from '../../graph/graph.people';
import { IDynamicPerson, ViewType } from '../../graph/types';
import { getPersonImage } from '../../graph/graph.photos';
import { getUserWithPhoto } from '../../graph/graph.userWithPhoto';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';
import { getUserPresence } from '../../graph/graph.presence';
import { getPersonCardGraphData } from './mgt-person-card.graph';
import { MgtPerson } from '../mgt-person/mgt-person';
import { styles } from './mgt-person-card-css';
import { BasePersonCardSection } from './sections/BasePersonCardSection';
import { MgtPersonCardContact } from './sections/mgt-person-card-contact/mgt-person-card-contact';
import { MgtPersonCardFiles } from './sections/mgt-person-card-files/mgt-person-card-files';
import { MgtPersonCardMessages } from './sections/mgt-person-card-messages/mgt-person-card-messages';
import { MgtPersonCardOrganization } from './sections/mgt-person-card-organization/mgt-person-card-organization';
import { MgtPersonCardProfile } from './sections/mgt-person-card-profile/mgt-person-card-profile';
import { MgtPersonCardConfig, MgtPersonCardState } from './mgt-person-card.types';
import { strings } from './strings';

import '../sub-components/mgt-spinner/mgt-spinner';

export * from './mgt-person-card.types';

// tslint:disable-next-line:completed-docs
interface MgtPersonCardStateHistory {
  // tslint:disable-next-line:completed-docs
  state: MgtPersonCardState;
  // tslint:disable-next-line:completed-docs
  personDetails: IDynamicPerson;
  // tslint:disable-next-line:completed-docs
  personImage: string;
}

/**
 * Web Component used to show detailed data for a person in the Microsoft Graph
 *
 * @export
 * @class MgtPersonCard
 * @extends {MgtTemplatedComponent}
 *
 * @fires expanded - Fired when expanded details section is opened
 *
 * @cssprop --person-card-display-name-font-size - {Length} Font size of display name title
 * @cssprop --person-card-display-name-color - {Color} Color of display name font
 * @cssprop --person-card-title-font-size - {Length} Font size of title
 * @cssprop --person-card-title-color - {Color} Color of title
 * @cssprop --person-card-subtitle-font-size - {Length} Font size of subtitle
 * @cssprop --person-card-subtitle-color - {Color} Color of subttitle
 * @cssprop --person-card-details-title-font-size - {Length} Font size additional details title
 * @cssprop --person-card-details-title-color- {Color} Color of additional details title
 * @cssprop --person-card-details-item-font-size - {Length} Font size items in additional details section
 * @cssprop --person-card-details-item-color - {Color} Color of items in additional details section
 * @cssprop --person-card-background-color - {Color} Color of person card background
 */
@customElement('mgt-person-card')
export class MgtPersonCard extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  protected get strings() {
    return strings;
  }

  /**
   * Get the scopes required for the person card
   * The scopes depend on what sections are shown
   *
   * Use the `MgtPersonCard.config` object to configure
   * what sections are shown
   *
   * @static
   * @return {*}  {string[]}
   * @memberof MgtPersonCard
   */
  public static get requiredScopes(): string[] {
    return MgtPersonCard.getScopes();
  }

  public static getScopes(): string[] {
    const scopes = [];

    if (this.config.sections.files) {
      scopes.push('Sites.Read.All');
    }

    if (this.config.sections.mailMessages) {
      scopes.push('Mail.ReadBasic');
    }

    if (this.config.sections.organization) {
      scopes.push('User.Read.All');

      if (typeof this.config.sections.organization !== 'boolean' && this.config.sections.organization.showWorksWith) {
        scopes.push('People.Read.All');
      }
    }

    if (this.config.sections.profile) {
      scopes.push('User.Read.All');
    }

    if (this.config.useContactApis) {
      scopes.push('Contacts.Read');
    }

    if (scopes.indexOf('User.Read.All') < 0) {
      // at minimum, we need these scopes
      scopes.push('User.ReadBasic.All');
      scopes.push('User.Read');
    }

    if (scopes.indexOf('People.Read.All') < 0) {
      // at minimum, we need these scopes
      scopes.push('People.Read');
    }

    // return unique
    return [...new Set(scopes)];
  }

  /**
   * Global configuration object for
   * all person card components
   *
   * @static
   * @type {MgtPersonCardConfig}
   * @memberof MgtPersonCard
   */
  public static get config() {
    return this._config;
  }

  private static _config: MgtPersonCardConfig = {
    sections: {
      files: true,
      mailMessages: true,
      organization: { showWorksWith: true },
      profile: true
    },
    useContactApis: true
  };

  /**
   * Set the person details to render
   *
   * @type {IDynamicPerson}
   * @memberof MgtPersonCard
   */
  @property({
    attribute: 'person-details',
    type: Object
  })
  public get personDetails(): IDynamicPerson {
    return this._personDetails;
  }
  public set personDetails(value: IDynamicPerson) {
    if (this._personDetails === value) {
      return;
    }

    this._personDetails = value;
    this.personImage = this.getImage();
    this.requestStateUpdate();
  }
  /**
   * allows developer to define name of person for component
   * @type {string}
   */
  @property({
    attribute: 'person-query'
  })
  public personQuery: string;

  /**
   * allows the locking of navigation using tabs to not flow out of the card section
   * @type {boolean}
   */
  @property({
    attribute: 'lock-tab-navigation',
    type: Boolean
  })
  public lockTabNavigation: boolean;

  /**
   * user-id property allows developer to use id value for component
   * @type {string}
   */
  @property({
    attribute: 'user-id'
  })
  public get userId(): string {
    return this._userId;
  }
  public set userId(value: string) {
    if (value === this._userId) {
      return;
    }
    this._userId = value;
    this.personDetails = null;
    this.state = null;
    this.requestStateUpdate();
  }

  /**
   * Set the image of the person
   * Set to '@' to look up image from the graph
   *
   * @type {string}
   * @memberof MgtPersonCard
   */
  @property({
    attribute: 'person-image',
    type: String
  })
  public personImage: string;

  /**
   * Sets whether the person image should be fetched
   * from the Microsoft Graph based on the personDetails
   * provided by the user
   *
   * @type {boolean}
   * @memberof MgtPerson
   */
  @property({
    attribute: 'fetch-image',
    type: Boolean
  })
  public fetchImage: boolean;

  /**
   * Gets or sets whether expanded details section is rendered
   *
   * @type {boolean}
   * @memberof MgtPersonCard
   */
  @property({
    attribute: 'is-expanded',
    type: Boolean
  })
  public isExpanded: boolean;

  /**
   * Gets or sets whether person details should be inherited from an mgt-person parent
   * Useful when used as template in an mgt-person component
   *
   * @type {boolean}
   * @memberof MgtPersonCard
   */
  @property({
    attribute: 'inherit-details',
    type: Boolean
  })
  public inheritDetails: boolean;

  /**
   * determines if person card component renders presence
   * @type {boolean}
   */
  @property({
    attribute: 'show-presence',
    type: Boolean
  })
  public showPresence: boolean;

  /**
   * Gets or sets presence of person
   *
   * @type {MicrosoftGraph.Presence}
   * @memberof MgtPerson
   */
  @property({
    attribute: 'person-presence',
    type: Object
  })
  public personPresence: Presence;

  /**
   * The subsections for display in the lower part of the card
   *
   * @protected
   * @type {BasePersonCardSection[]}
   * @memberof MgtPersonCard
   */
  protected sections: BasePersonCardSection[];

  @internalProperty() private state: MgtPersonCardState;
  @internalProperty() private isStateLoading: boolean;

  private _history: MgtPersonCardStateHistory[];
  private _chatInput: string;
  private _currentSection: BasePersonCardSection;
  private _personDetails: IDynamicPerson;
  private _me: User;
  private _smallView;
  private _windowHeight;

  private _userId: string;

  private get internalPersonDetails(): IDynamicPerson {
    return (this.state && this.state.person) || this.personDetails;
  }

  constructor() {
    super();
    this._chatInput = '';
    this._currentSection = null;
    this._history = [];
    this.sections = [];
  }

  /**
   * Synchronizes property values when attributes change.
   *
   * @param {*} name
   * @param {*} oldValue
   * @param {*} newValue
   * @memberof MgtPersonCard
   */
  public attributeChangedCallback(name: string, oldValue: string, newValue: string) {
    super.attributeChangedCallback(name, oldValue, newValue);

    if (oldValue === newValue) {
      return;
    }

    switch (name) {
      case 'person-query':
        this.personDetails = null;
        this.state = null;
        this.requestStateUpdate();
        break;
    }
  }

  /**
   * Navigate the card to a different person.
   *
   * @protected
   * @memberof MgtPersonCard
   */
  public navigate(person: IDynamicPerson): void {
    this._history.push({
      personDetails: this.personDetails,
      personImage: this.getImage(),
      state: this.state
    });

    this._personDetails = person;
    this.state = null;
    this.personImage = null;
    this._currentSection = null;
    this.sections = [];
    this.requestStateUpdate();
  }

  /**
   * Navigate the card back to the previous person
   *
   * @returns {void}
   * @memberof MgtPersonCard
   */
  public goBack(): void {
    if (!this._history || !this._history.length) {
      return;
    }

    const historyState = this._history.pop();
    this._currentSection = null;

    this.state = historyState.state;
    this._personDetails = historyState.state;
    this.personImage = historyState.personImage;
    this.loadSections();
  }

  /**
   * Navigate the card back to first person and clear history
   *
   * @returns {void}
   * @memberof MgtPersonCard
   */
  public clearHistory(): void {
    this._currentSection = null;

    if (!this._history || !this._history.length) {
      return;
    }

    const historyState = this._history[0];
    this._history = [];

    this.state = historyState.state;
    this._personDetails = historyState.state;
    this.personImage = historyState.personImage;
    this.loadSections();
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    // Handle no data
    if (!this.internalPersonDetails) {
      return this.renderNoData();
    }

    const person = this.internalPersonDetails;
    const image = this.getImage();

    // Check for a default template.
    // tslint:disable-next-line: no-string-literal
    if (this.hasTemplate('default')) {
      return this.renderTemplate('default', {
        person: this.internalPersonDetails,
        personImage: image
      });
    }

    const navigationTemplate =
      this._history && this._history.length
        ? html`
            <div class="nav">
              <div class="nav__back" tabindex="0" @keydown=${(e: KeyboardEvent) => {
                e.code === 'Enter' ? this.goBack() : '';
              }} @click=${() => this.goBack()}>${getSvg(SvgIcon.Back)}</div>
            </div>
          `
        : null;

    // Check for a person-details template
    let personDetailsTemplate = this.renderTemplate('person-details', {
      person: this.internalPersonDetails,
      personImage: image
    });
    if (!personDetailsTemplate) {
      const personTemplate = this.renderPerson();
      const contactIconsTemplate = this.renderContactIcons(person);

      personDetailsTemplate = html`
        ${personTemplate} ${contactIconsTemplate}
      `;
    }

    const expandedDetailsTemplate = this.isExpanded ? this.renderExpandedDetails() : this.renderExpandedDetailsButton();
    this._windowHeight =
      window.innerHeight && document.documentElement.clientHeight
        ? Math.min(window.innerHeight, document.documentElement.clientHeight)
        : window.innerHeight || document.documentElement.clientHeight;

    if (this._windowHeight < 250) {
      this._smallView = true;
    }
    const tabLocker = this.lockTabNavigation
      ? html`<div @keydown=${this.handleEndOfCard} aria-label=${this.strings.endOfCard} tabindex="0" id="end-of-container"></div>`
      : html``;
    return html`
      <div class="root" dir=${this.direction}>
      <div class=${this._smallView ? 'small' : ''}>
          ${navigationTemplate}
          <div class="person-details-container">${personDetailsTemplate}</div>
          <div class="expanded-details-container">${expandedDetailsTemplate}</div>
          ${tabLocker}
        </div>
      </div>
    `;
  }

  private handleEndOfCard(e: KeyboardEvent) {
    if (e && e.code === 'Tab') {
      const endOfCardEl = this.renderRoot.querySelector('#end-of-container') as HTMLElement;
      if (endOfCardEl) {
        endOfCardEl.blur();
        const imageCardEl = this.renderRoot.querySelector('mgt-person') as HTMLElement;
        if (imageCardEl) {
          imageCardEl.focus();
        }
      }
    }
  }

  /**
   * Render the state when no data is available.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCard
   */
  protected renderNoData(): TemplateResult {
    return this.renderTemplate('no-data', null) || html``;
  }

  /**
   * Render a display image for the person.
   *
   * @protected
   * @param {*} image
   * @memberof MgtPersonCard
   */
  protected renderPerson(): TemplateResult {
    const avatarSize = 'large';
    return html`
      <mgt-person
        tabindex="0"
        class="person-image"
        .personDetails=${this.internalPersonDetails}
        .personImage=${this.getImage()}
        .personPresence=${this.personPresence}
        .showPresence=${this.showPresence}
        .avatarSize=${avatarSize}
        .view=${ViewType.threelines}
        .line2Property=${'jobTitle'}
        .line3Property=${'department'}
      ></mgt-person>
    `;
  }

  /**
   * Render person subtitle.
   *
   * @protected
   * @param {IDynamicPerson} person
   * @returns {TemplateResult}
   * @memberof MgtPersonCard
   */
  protected renderPersonSubtitle(person?: IDynamicPerson): TemplateResult {
    person = person || this.internalPersonDetails;
    if (!person.department) {
      return;
    }
    return html`
      <div class="department">${person.department}</div>
    `;
  }

  /**
   * Render the various icons for contacting the person.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCard
   */
  protected renderContactIcons(person?: IDynamicPerson): TemplateResult {
    if (this.isExpanded) {
      return;
    }

    person = person || this.internalPersonDetails;
    const userPerson = person as User;

    // Email
    let email: TemplateResult;
    if (getEmailFromGraphEntity(person)) {
      email = html`
        <div class="icon" @click=${() => this.emailUser()} tabindex=0>
          ${getSvg(SvgIcon.SmallEmail)}
          <span>${this.strings.sendEmailLinkSubtitle}</span>
        </div>
      `;
    }

    // Chat
    let chat: TemplateResult;
    if (userPerson.userPrincipalName) {
      chat = html`
        <div class="icon" @click=${() => this.chatUser()} tabindex=0>
          ${getSvg(SvgIcon.SmallChat)}
          <span>${this.strings.startChatLinkSubtitle}</span>
        </div>
      `;
    }

    return html`
      <div class="base-icons">
        ${email} ${chat}
      </div>
    `;
  }

  /**
   * Render the button used to expand the expanded details.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCard
   */
  protected renderExpandedDetailsButton(): TemplateResult {
    return html`
      <div 
        class="expanded-details-button" 
        @click=${() => this.showExpandedDetails()} 
        @keydown=${this.handleKeyDown}
        tabindex=0>
          ${getSvg(SvgIcon.ExpandDown)}
      </div>
    `;
  }

  /**
   * Render expanded details.
   *
   * @protected
   * @param {IDynamicPerson} [person]
   * @returns {TemplateResult}
   * @memberof MgtPersonCard
   */
  protected renderExpandedDetails(person?: IDynamicPerson): TemplateResult {
    if (!this.state && this.isStateLoading) {
      return html`
        <div class="loading">
          <mgt-spinner></mgt-spinner>
        </div>
      `;
    }
    // load sections when details are expanded
    // when not singed in
    const provider = Providers.globalProvider;
    if (provider.state === ProviderState.SignedOut) {
      this.loadSections();
    }

    person = person || this.internalPersonDetails;

    const sectionNavTemplate = this.renderSectionNavigation();
    const currentSectionTemplate = this.renderCurrentSection();

    return html`
      <div class="section-nav">
        ${sectionNavTemplate}
      </div>
      <div class="section-host ${this._smallView ? 'small' : ''}" @wheel=${(e: WheelEvent) =>
      this.handleSectionScroll(e)} tabindex=0>
        ${currentSectionTemplate}
      </div>
    `;
  }

  /**
   * Render the navigation ribbon for subsections
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCard
   */
  protected renderSectionNavigation(): TemplateResult {
    if (!this.sections || (this.sections.length < 2 && !this.hasTemplate('additional-details'))) {
      return;
    }

    const currentSectionIndex = this._currentSection ? this.sections.indexOf(this._currentSection) : -1;

    const navIcons = this.sections.map((section, i) => {
      const classes = classMap({
        active: i === currentSectionIndex,
        'section-nav__icon': true
      });
      const tagName = section.tagName;
      const ariaLabel = tagName.substring(16, tagName.length).toLowerCase();
      return html`
        <button
          aria-label=${ariaLabel}
          tabindex=0
          class=${classes}
          @click=${() => this.updateCurrentSection(section)}>
            ${section.renderIcon()}</button>
      `;
    });

    const overviewClasses = classMap({
      active: currentSectionIndex === -1,
      'section-nav__icon': true
    });
    return html`
      <button 
        aria-label="overview"
        tabindex=0 
        class=${overviewClasses}
        @click=${() => this.updateCurrentSection(null)}>
          ${getSvg(SvgIcon.Overview)}
      </button>
      ${navIcons}
    `;
  }

  /**
   * Render the default section with compact views for each subsection.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCard
   */
  protected renderOverviewSection(): TemplateResult {
    function handleKeyDown(e: KeyboardEvent, section: BasePersonCardSection) {
      e.code === 'Enter' ? this.updateCurrentSection(section) : '';
    }

    const compactTemplates = this.sections.map(
      (section: BasePersonCardSection) => html`
        <div class="section">
          <div class="section__header">
            <div class="section__title" tabindex="0">${section.displayName}</div>
            <a 
              class="section__show-more"
              tabindex=0
              @keydown=${(e: KeyboardEvent) => handleKeyDown(e, section)}
              @click=${() => this.updateCurrentSection(section)}>
                ${this.strings.showMoreSectionButton}</a>
          </div>
          <div class="section__content">${section.asCompactView()}</div>
        </div>
      `
    );

    const additionalDetails = this.renderTemplate('additional-details', {
      person: this.internalPersonDetails,
      personImage: this.getImage(),
      state: this.state
    });
    if (additionalDetails) {
      compactTemplates.splice(
        1,
        0,
        html`
          <div class="section">
            <div class="additional-details">${additionalDetails}</div>
          </div>
        `
      );
    }

    // To be included when support for direct messaging is added
    // ${this.internalPersonDetails.id !== this._me.id && MgtPersonCard.config.isSendMessageVisible
    //   ? html`
    //       <div class="quick-message">
    //         <input
    //           type="text"
    //           class="quick-message__input"
    //           placeholder="Message ${this.internalPersonDetails.displayName}"
    //           .value=${this._chatInput}
    //           @input=${(e: Event) => {
    //             this._chatInput = (e.target as HTMLInputElement).value;
    //           }}
    //         />
    //         <button class="quick-message__send" @click=${() => this.sendQuickMessage()}>
    //           ${getSvg(SvgIcon.Send)}
    //         </button>
    //       </div>
    //     `
    //   : null}

    return html`
      <div class="sections">
        ${compactTemplates}
      </div>
    `;
  }

  /**
   * Render the actively selected section.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCard
   */
  protected renderCurrentSection(): TemplateResult {
    if ((!this.sections || !this.sections.length) && !this.hasTemplate('additional-details')) {
      return;
    }

    if (this.sections.length === 1 && !this.hasTemplate('additional-details')) {
      return html`
        ${this.sections[0].asFullView()}
      `;
    }

    if (!this._currentSection) {
      return this.renderOverviewSection();
    }

    return html`
      ${this._currentSection.asFullView()}
    `;
  }

  /**
   * load state into the component
   *
   * @protected
   * @returns
   * @memberof MgtPersonCard
   */
  protected async loadState() {
    if (this.state) {
      return;
    }

    if (!this.personDetails && this.inheritDetails) {
      // User person details inherited from parent tree
      let parent = this.parentElement;
      while (parent && parent.tagName !== 'MGT-PERSON') {
        parent = parent.parentElement;
      }

      let parentPerson = (parent as MgtPerson).personDetails || parent['personDetailsInternal'];

      if (parent && parentPerson) {
        this.personDetails = parentPerson;
        this.personImage = (parent as MgtPerson).personImage;
      }
    }

    const provider = Providers.globalProvider;

    // check if user is signed in
    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    const graph = provider.graph.forComponent(this);

    this.isStateLoading = true;

    if (!this._me) {
      this._me = await Providers.me();
    }

    // check if personDetail already populated
    if (this.personDetails) {
      const user = this.personDetails as User;
      const id = user.userPrincipalName || user.id;

      // if we have an id but no email, we should get data from the graph
      // in some graph calls, the user object does not contain the email
      if (id && !getEmailFromGraphEntity(user)) {
        const person = await getUserWithPhoto(graph, id);
        this.personDetails = person;
        this.personImage = this.getImage();
      }
    } else if (this.userId || this.personQuery === 'me') {
      // Use userId or 'me' query to get the person and image
      const person = await getUserWithPhoto(graph, this.userId);
      this.personDetails = person;
      this.personImage = this.getImage();
    } else if (this.personQuery) {
      // Use the personQuery to find our person.
      const people = await findPeople(graph, this.personQuery, 1);

      if (people && people.length) {
        this.personDetails = people[0];
        getPersonImage(graph, this.personDetails, MgtPersonCard.config.useContactApis).then(image => {
          if (image) {
            this.personDetails.personImage = image;
            this.personImage = image;
          }
        });
      }
    }

    // populate presence
    const defaultPresence = {
      activity: 'Offline',
      availability: 'Offline',
      id: null
    };
    if (!this.personPresence && this.showPresence) {
      try {
        if (this.personDetails && this.personDetails.id) {
          getUserPresence(graph, this.personDetails.id).then(presence => {
            this.personPresence = presence;
          });
        } else {
          this.personPresence = defaultPresence;
        }
      } catch (_) {
        // set up a default Presence in case beta api changes or getting error code
        this.personPresence = defaultPresence;
      }
    }

    // populate state
    if (this.personDetails?.id) {
      this.state = await getPersonCardGraphData(
        graph,
        this.personDetails,
        this._me === this.personDetails.id,
        MgtPersonCard.config
      );
    }

    this.loadSections();

    this.isStateLoading = false;
  }

  /**
   * Send a chat message to the user from the quick message input.
   *
   * @protected
   * @returns {void}
   * @memberof MgtPersonCard
   */
  protected sendQuickMessage(): void {
    const message = this._chatInput.trim();
    if (!message || !message.length) {
      return;
    }

    this.chatUser(message);
  }

  /**
   * Use the mailto: protocol to initiate a new email to the user.
   *
   * @protected
   * @memberof MgtPersonCard
   */
  protected emailUser() {
    const user = this.internalPersonDetails;
    if (user) {
      const email = getEmailFromGraphEntity(user);
      if (email) {
        window.open('mailto:' + email, '_blank', 'noreferrer');
      }
    }
  }

  /**
   * Use the tel: protocol to initiate a new call to the user.
   *
   * @protected
   * @memberof MgtPersonCard
   */
  protected callUser() {
    const user = this.personDetails as User;
    const person = this.personDetails as microsoftgraph.Person;

    if (user && user.businessPhones && user.businessPhones.length) {
      const phone = user.businessPhones[0];
      if (phone) {
        window.open('tel:' + phone, '_blank', 'noreferrer');
      }
    } else if (person && person.phones && person.phones.length) {
      const businessPhones = this.getPersonBusinessPhones(person);
      const phone = businessPhones[0];
      if (phone) {
        window.open('tel:' + phone, '_blank', 'noreferrer');
      }
    }
  }

  /**
   * Initiate a chat message to the user via deeplink.
   *
   * @protected
   * @memberof MgtPersonCard
   */
  protected chatUser(message: string = null) {
    const user = this.personDetails as User;
    if (user && user.userPrincipalName) {
      const users: string = user.userPrincipalName;

      let url = `https://teams.microsoft.com/l/chat/0/0?users=${users}`;
      if (message && message.length) {
        url += `&message=${message}`;
      }

      const openWindow = () => window.open(url, '_blank', 'noreferrer');

      if (TeamsHelper.isAvailable) {
        TeamsHelper.executeDeepLink(url, (status: boolean) => {
          if (!status) {
            openWindow();
          }
        });
      } else {
        openWindow();
      }
    }
  }

  /**
   * Display the expanded details panel.
   *
   * @protected
   * @memberof MgtPersonCard
   */
  protected showExpandedDetails() {
    const root = this.renderRoot.querySelector('.root');
    if (root && root.animate) {
      // play back
      root.animate(
        [
          {
            height: 'auto',
            transformOrigin: 'top left'
          },
          {
            height: 'auto',
            transformOrigin: 'top left'
          }
        ],
        {
          duration: 1000,
          easing: 'ease-in-out',
          fill: 'both'
        }
      );
    }
    this.isExpanded = true;

    this.fireCustomEvent('expanded', null, true);
  }

  private loadSections() {
    this.sections = [];

    if (!this.internalPersonDetails) {
      return;
    }

    const contactSections = new MgtPersonCardContact(this.internalPersonDetails as User);
    if (contactSections.hasData) {
      this.sections.push(contactSections);
    }

    if (!this.state) {
      return;
    }

    const { person, directReports, messages, files, profile } = this.state;

    if (
      MgtPersonCard.config.sections.organization &&
      ((person && person.manager) || (directReports && directReports.length))
    ) {
      this.sections.push(new MgtPersonCardOrganization(this.state, this._me));
    }

    if (MgtPersonCard.config.sections.mailMessages && messages && messages.length) {
      this.sections.push(new MgtPersonCardMessages(messages));
    }

    if (MgtPersonCard.config.sections.files && files && files.length) {
      this.sections.push(new MgtPersonCardFiles(files));
    }

    if (MgtPersonCard.config.sections.profile && profile) {
      const profileSection = new MgtPersonCardProfile(profile);
      if (profileSection.hasData) {
        this.sections.push(profileSection);
      }
    }
  }

  private getImage(): string {
    if (this.personImage) {
      return this.personImage;
    }

    const person = this.personDetails;
    return person && person.personImage ? person.personImage : null;
  }

  private getPersonBusinessPhones(person: Person): string[] {
    const phones = person.phones;
    const businessPhones: string[] = [];
    for (const p of phones) {
      if (p.type === 'business') {
        businessPhones.push(p.number);
      }
    }
    return businessPhones;
  }

  private updateCurrentSection(section) {
    const sectionHost = this.renderRoot.querySelector('.section-host');
    sectionHost.scrollTop = 0;

    this._currentSection = section;
    this.requestUpdate();
  }

  private handleSectionScroll(e: WheelEvent) {
    const target = this.renderRoot.querySelector('.section-host') as HTMLElement;
    if (target) {
      if (
        !(e.deltaY < 0 && target.scrollTop === 0) &&
        !(e.deltaY > 0 && target.clientHeight + target.scrollTop >= target.scrollHeight - 1)
      ) {
        e.stopPropagation();
      }
    }
  }

  private handleKeyDown(e: KeyboardEvent) {
    //enter activates person-card
    if (e) {
      if (e.keyCode === 13) {
        this.showExpandedDetails();
      }
    }
  }
}
