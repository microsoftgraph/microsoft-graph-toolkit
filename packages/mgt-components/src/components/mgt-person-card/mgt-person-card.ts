/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html, TemplateResult } from 'lit';
import { property, state } from 'lit/decorators.js';
import { classMap } from 'lit/directives/class-map.js';
import {
  MgtTemplatedComponent,
  Providers,
  ProviderState,
  TeamsHelper,
  mgtHtml,
  customElement,
  customElementHelper
} from '@microsoft/mgt-element';
import { IGraph } from '@microsoft/mgt-element';
import { Presence, User, Person } from '@microsoft/microsoft-graph-types';

import { findPeople, getEmailFromGraphEntity } from '../../graph/graph.people';
import { IDynamicPerson, ViewType } from '../../graph/types';
import { getPersonImage } from '../../graph/graph.photos';
import { getUserWithPhoto } from '../../graph/graph.userWithPhoto';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';
import { getUserPresence } from '../../graph/graph.presence';
import { getPersonCardGraphData, createChat, sendMessage } from './mgt-person-card.graph';
import { MgtPerson } from '../mgt-person/mgt-person';
import { styles } from './mgt-person-card-css';
import { MgtContact } from '../mgt-contact/mgt-contact';
import { MgtFileList } from '../mgt-file-list/mgt-file-list';
import { MgtMessages } from '../mgt-messages/mgt-messages';
import { MgtOrganization } from '../mgt-organization/mgt-organization';
import { MgtProfile } from '../mgt-profile/mgt-profile';
import { MgtPersonCardConfig, MgtPersonCardState } from './mgt-person-card.types';
import { strings } from './strings';
import { isUser } from '../../graph/entityType';

import '../sub-components/mgt-spinner/mgt-spinner';

export * from './mgt-person-card.types';

import {
  fluentTabs,
  fluentTab,
  fluentTabPanel,
  fluentButton,
  fluentTextField,
  fluentCard
} from '@fluentui/web-components';
import { registerFluentComponents } from '../../utils/FluentComponents';
import { BasePersonCardSection, CardSection } from '../BasePersonCardSection';

registerFluentComponents(fluentCard, fluentTabs, fluentTab, fluentTabPanel, fluentButton, fluentTextField);

interface MgtPersonCardStateHistory {
  state: MgtPersonCardState;
  personDetails: IDynamicPerson;
  personImage: string;
}

/**
 * Web Component used to show detailed data for a person in the Microsoft Graph
 *
 * @export
 * @class MgtPersonCard
 * @extends {MgtTemplatedComponent}
 *
 * @fires {CustomEvent<null>} expanded - Fired when expanded details section is opened
 *
 * @cssprop --person-card-line1-font-size - {Length} Font size of line 1
 * @cssprop --person-card-line1-font-weight - {FontWeight} Font weight of line 1
 * @cssprop --person-card-line1-line-height - {Length} Line height of line 1
 * @cssprop --person-card-line2-font-size - {Length} Font size of line 2
 * @cssprop --person-card-line2-font-weight - {FontWeight} Font weight of line 2
 * @cssprop --person-card-line2-line-height - {Length} Line height of line 2
 * @cssprop --person-card-line3-font-size - {Length} Font size of line 3
 * @cssprop --person-card-line3-font-weight - {FontWeight} Font weight of line 3
 * @cssprop --person-card-line3-line-height - {Length} Line height of line 3
 * @cssprop --person-card-avatar-size - {Length} Width and height of the avatar. Default is 75px
 * @cssprop --person-card-details-left-spacing - {Length} The space between the avatar and the person details. Default is 15px
 * @cssprop --person-card-avatar-top-spacing - {Length} The margin top of the avatar in person-card component
 * @cssprop --person-card-details-bottom-spacing - {Length} The margin bottom of the person details in person-card component
 * @cssprop --person-card-base-icons-left-spacing - {Length} The margin-inline-start of the base-icons in person-card component
 * @cssprop --person-card-background-color - {Color} The color of the person-card-component
 * @cssprop --person-card-expanded-background-color-hover - {Color} The hover color of the expanded details button of the person card component
 * @cssprop --person-card-nav-back-arrow-hover-color - {Color} The hover color of the back arrow of the person card component
 * @cssprop --person-card-icon-color - {Color} The color of the icons in the person card component
 * @cssprop --person-card-icon-hover-color - {Color} The hover color of the icons in the person card component
 * @cssprop --person-card-show-more-color - {Color} The color of the show more text in the person card component
 * @cssprop --person-card-show-more-hover-color - {Color} The hover color of the show more text in person card component
 * @cssprop --person-card-fluent-background-color - {Color} The background color of the fluent buttons in person card component
 * @cssprop --person-card-line1-text-color - {Color} The color of line 1 in person card
 * @cssprop --person-card-line2-text-color - {Color} The color of line 2 in person card
 * @cssprop --person-card-line3-text-color - {Color} The color of line 3 in person card
 * @cssprop --person-card-fluent-background-color-hover - {Color} The hover background color of the fluent buttons in person card component
 * @cssprop --person-card-chat-input-hover-color - {Color} The chat input hover color
 * @cssprop --person-card-chat-input-focus-color - {Color} The chat input focus color
 */
@customElement('person-card')
export class MgtPersonCard extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * Strings to use for localization
   *
   * @readonly
   * @protected
   * @memberof MgtPersonCard
   */
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

  /**
   * Scopes used to fetch data for the component
   *
   * @static
   * @return {*}  {string[]}
   * @memberof MgtPersonCard
   */
  public static getScopes(): string[] {
    const scopes: string[] = [];

    if (this.config.sections.files) {
      scopes.push('Sites.Read.All');
    }

    if (this.config.sections.mailMessages) {
      scopes.push('Mail.Read');
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

    scopes.push('Chat.Create', 'Chat.ReadWrite');

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

  private static readonly _config: MgtPersonCardConfig = {
    sections: {
      files: true,
      mailMessages: true,
      organization: { showWorksWith: true },
      profile: true
    },
    useContactApis: true,
    isSendMessageVisible: true
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
    void this.requestStateUpdate();
  }
  /**
   * allows developer to define name of person for component
   *
   * @type {string}
   */
  @property({
    attribute: 'person-query'
  })
  public personQuery: string;

  /**
   * allows the locking of navigation using tabs to not flow out of the card section
   *
   * @type {boolean}
   */
  @property({
    attribute: 'lock-tab-navigation',
    type: Boolean
  })
  public lockTabNavigation: boolean;

  /**
   * user-id property allows developer to use id value for component
   *
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
    this._cardState = null;
    void this.requestStateUpdate();
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
   *
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

  @state()
  private isSending = false;

  /**
   * The subsections for display in the lower part of the card
   *
   * @protected
   * @type {any[]}
   * @memberof MgtPersonCard
   */
  protected sections: CardSection[];

  @state() private _cardState: MgtPersonCardState;
  @state() private _isStateLoading: boolean;

  private _history: MgtPersonCardStateHistory[];
  private _chatInput: string;
  private _currentSection: CardSection;
  private _personDetails: IDynamicPerson;
  private _me: User;
  private _smallView: boolean;
  private _windowHeight;

  private _userId: string;
  private _graph: IGraph;

  // eslint-disable-next-line @typescript-eslint/naming-convention
  private get internalPersonDetails(): IDynamicPerson {
    return this._cardState?.person || this.personDetails;
  }

  constructor() {
    super();
    this._chatInput = '';
    this._currentSection = null;
    this._history = [];
    this.sections = [];
    this._graph = null;
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
        this._cardState = null;
        void this.requestStateUpdate();
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
      state: this._cardState
    });

    this._personDetails = person;
    this._cardState = null;
    this.personImage = null;
    this._currentSection = null;
    this.sections = [];
    this._chatInput = '';
    void this.requestStateUpdate();
  }

  /**
   * Navigate the card back to the previous person
   *
   * @returns {void}
   * @memberof MgtPersonCard
   */
  public goBack = (): void => {
    if (!this._history?.length) {
      return;
    }

    const historyState = this._history.pop();
    this._currentSection = null;

    // resets to first tab being selected
    const firstTab: HTMLElement = this.renderRoot.querySelector('fluent-tab');
    if (firstTab) {
      firstTab.click();
    }
    this._cardState = historyState.state;
    this._personDetails = historyState.state.person;
    this.personImage = historyState.personImage;
    this.loadSections();
  };

  /**
   * Navigate the card back to first person and clear history
   *
   * @returns {void}
   * @memberof MgtPersonCard
   */
  public clearHistory(): void {
    this._currentSection = null;

    if (!this._history?.length) {
      return;
    }

    const historyState = this._history[0];
    this._history = [];

    this._cardState = historyState.state;
    this._personDetails = historyState.personDetails;
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
    // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access, @typescript-eslint/dot-notation
    if (this.hasTemplate('default')) {
      return this.renderTemplate('default', {
        person: this.internalPersonDetails,
        personImage: image
      });
    }

    let ariaLabel: string;

    ariaLabel = this.strings.closeCardLabel;

    const closeCardTemplate = this.isExpanded
      ? html`
           <div class="close-card-container">
             <fluent-button 
              appearance="lightweight" 
              class="close-button" 
              aria-label=${ariaLabel}
              @click=${this.closeCard} >
               ${getSvg(SvgIcon.Close)}
             </fluent-button>
           </div>
         `
      : null;

    ariaLabel = this.strings.goBackLabel;
    const navigationTemplate = this._history?.length
      ? html`
            <div class="nav">
              <fluent-button 
                appearance="lightweight"
                class="nav__back" 
                aria-label=${ariaLabel} 
                @keydown=${this.handleGoBack}
                @click=${this.goBack}>${getSvg(SvgIcon.Back)}
               </fluent-button>
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
        <div class=${classMap({ small: this._smallView })}>
          ${navigationTemplate}
          ${closeCardTemplate}
          <div class="person-details-container">${personDetailsTemplate}</div>
          <div class="expanded-details-container">${expandedDetailsTemplate}</div>
          ${tabLocker}
        </div>
      </div>
     `;
  }

  private readonly handleEndOfCard = (e: KeyboardEvent) => {
    if (e && e.code === 'Tab') {
      const endOfCardEl = this.renderRoot.querySelector<HTMLElement>('#end-of-container');
      if (endOfCardEl) {
        endOfCardEl.blur();
        const imageCardEl = this.renderRoot.querySelector<HTMLElement>('mgt-person');
        if (imageCardEl) {
          imageCardEl.focus();
        }
      }
    }
  };

  /**
   * Render the state when no data is available.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCard
   */
  protected closeCard = () => {
    // reset tabs
    this.updateCurrentSection(null);
    this.isExpanded = false;
  };

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
    return mgtHtml`
      <mgt-person
        tabindex="0"
        class="person-image"
        .personDetails=${this.internalPersonDetails}
        .personImage=${this.getImage()}
        .personPresence=${this.personPresence}
        .showPresence=${this.showPresence}
        .view=${ViewType.threelines}
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
    if (!isUser(person) || !person.department) {
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
    person = person || this.internalPersonDetails;
    const userPerson = person as User;

    let ariaLabel: string;

    // Email
    let email: TemplateResult;
    if (getEmailFromGraphEntity(person)) {
      ariaLabel = `${this.strings.emailButtonLabel} ${person.displayName}`;
      email = html`
        <fluent-button class="icon"
          aria-label=${ariaLabel}
          @click=${this.emailUser}>
          ${getSvg(SvgIcon.SmallEmail)}
        </fluent-button>
      `;
    }

    // Chat
    let chat: TemplateResult;
    if (userPerson?.userPrincipalName) {
      ariaLabel = `${this.strings.chatButtonLabel} ${person.displayName}`;
      chat = html`
        <fluent-button class="icon"
          aria-label=${ariaLabel}
          @click=${this.chatUser}>
          ${getSvg(SvgIcon.SmallChat)}
        </fluent-button>
       `;
    }

    // Video

    ariaLabel = `${this.strings.videoButtonLabel} ${person.displayName}`;
    const video: TemplateResult = html`
      <fluent-button class="icon"
        aria-label=${ariaLabel}
        @click=${this.videoCallUser}>
        ${getSvg(SvgIcon.Video)}
      </fluent-button>
    `;

    // Call
    let call: TemplateResult;
    if (userPerson.userPrincipalName) {
      ariaLabel = `${this.strings.callButtonLabel} ${person.displayName}`;
      call = html`
         <fluent-button class="icon"
          aria-label=${ariaLabel}
          @click=${this.callUser}>
          ${getSvg(SvgIcon.Call)}
        </fluent-button>
       `;
    }

    return html`
       <div class="base-icons">
         ${email} ${chat} ${video} ${call}
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
      <fluent-button
        aria-label=${this.strings.expandDetailsLabel}
        class="expanded-details-button"
        @click=${this.showExpandedDetails}
      >
        ${getSvg(SvgIcon.ExpandDown)}
      </fluent-button>
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
  protected renderExpandedDetails(): TemplateResult {
    if (!this._cardState && this._isStateLoading) {
      return mgtHtml`
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

    const sectionNavTemplate = this.renderSectionNavigation();

    return html`
      <div class="section-nav">
        ${sectionNavTemplate}
      </div>
      <hr class="divider"/>
      <div
        class="section-host ${this._smallView ? 'small' : ''} ${this._smallView ? 'small' : ''}"
        @wheel=${(e: WheelEvent) => this.handleSectionScroll(e)}
        tabindex=0
      ></div>
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

    const additionalSectionTemplates = this.sections.map((section, i) => {
      const name = section.tagName.toLowerCase();
      const classes = classMap({
        active: i === currentSectionIndex,
        'section-nav__icon': true
      });

      return html`
        <fluent-tab
          id="${name}-Tab"
          class=${classes}
          slot="tab"
          @keyup="${() => this.updateCurrentSection(section)}"
          @click=${() => this.updateCurrentSection(section)}
        >
          ${section.renderIcon()}
        </fluent-tab>
      `;
    });

    const additionalPanelTemplates = this.sections.map(section => {
      return html`
        <fluent-tab-panel slot="tabpanel">
          <div class="inserted">
            <div class="title">${section.cardTitle}</div>
            ${this._currentSection ? section.asFullView() : null}
          </div>
        </fluent-tab-panel>
      `;
    });

    const overviewClasses = classMap({
      active: currentSectionIndex === -1,
      'section-nav__icon': true,
      overviewTab: true
    });

    return html`
      <fluent-tabs
        orientation="horizontal"
        activeindicator
        @wheel=${(e: WheelEvent) => this.handleSectionScroll(e)}
      >
        <fluent-tab
          class="${overviewClasses}"
          slot="tab"
          @keyup="${() => this.updateCurrentSection(null)}"
          @click=${() => this.updateCurrentSection(null)}
        >
          <div>${getSvg(SvgIcon.Overview)}</div>
        </fluent-tab>
        ${additionalSectionTemplates}
        <fluent-tab-panel slot="tabpanel" >
          <div class="overview-panel">${!this._currentSection ? this.renderOverviewSection() : null}</div>
        </fluent-tab-panel>
        ${additionalPanelTemplates}
      </fluent-tabs>
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
    const compactTemplates = this.sections.map(
      (section: BasePersonCardSection) => html`
        <div class="section">
          <div class="section__header">
            <div class="section__title" tabindex=0>${section.displayName}</div>
              <fluent-button
                appearance="lightweight"
                class="section__show-more"
                @click=${() => this.updateCurrentSection(section)}
              >
                ${this.strings.showMoreSectionButton}
              </fluent-button>
          </div>
          <div class="section__content">${section.asCompactView()}</div>
        </div>
      `
    );

    const additionalDetails = this.renderTemplate('additional-details', {
      person: this.internalPersonDetails,
      personImage: this.getImage(),
      state: this._cardState
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

    return html`
       <div class="sections">
          ${this.renderMessagingSection()}
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
    if (!this.sections?.length && !this.hasTemplate('additional-details')) {
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
   * Render the messaging section.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCard
   */
  protected renderMessagingSection(): TemplateResult {
    const person = this.personDetails as User;
    const user = this._me.userPrincipalName;
    const chatInput = this._chatInput;
    if (person?.userPrincipalName === user) {
      return;
    } else {
      return html`
      <div class="message-section">
        <fluent-text-field
          autocomplete="off"
          appearance="outline"
          placeholder="${this.strings.quickMessage}"
          .value=${chatInput}
          @input=${(e: Event) => {
            this._chatInput = (e.target as HTMLInputElement).value;
            this.requestUpdate();
          }}
          @keydown="${(e: KeyboardEvent) => this.sendQuickMessageOnEnter(e)}">
        </fluent-text-field>
        <fluent-button class="send-message-icon"
          aria-label=${this.strings.sendMessageLabel}
          @click=${this.sendQuickMessage}
          ?disabled=${this.isSending}>
          ${!this.isSending ? getSvg(SvgIcon.Send) : getSvg(SvgIcon.Confirmation)}
        </fluent-button>
      </div>
      `;
    }
  }

  /**
   * load state into the component
   *
   * @protected
   * @returns
   * @memberof MgtPersonCard
   */
  protected async loadState() {
    if (this._cardState) {
      return;
    }

    if (!this.personDetails && this.inheritDetails) {
      // User person details inherited from parent tree
      let parent = this.parentElement;
      while (parent && parent.tagName !== `${customElementHelper.prefix}-PERSON`.toUpperCase()) {
        parent = parent.parentElement;
      }

      const parentPerson: IDynamicPerson =
        // eslint-disable-next-line @typescript-eslint/dot-notation
        (parent as MgtPerson).personDetails || (parent as MgtPerson)['personDetailsInternal'];

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
    this._graph = graph;

    this._isStateLoading = true;

    if (!this._me) {
      this._me = await Providers.me();
    }

    // check if personDetail already populated
    if (this.personDetails) {
      const user = this.personDetails;
      let id: string;
      if (isUser(user)) {
        id = user.userPrincipalName || user.id;
      }

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

      if (people?.length) {
        this.personDetails = people[0];
        await getPersonImage(graph, this.personDetails, MgtPersonCard.config.useContactApis).then(image => {
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
        if (this.personDetails?.id) {
          this.personPresence = await getUserPresence(graph, this.personDetails.id);
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
      this._cardState = await getPersonCardGraphData(
        graph,
        this.personDetails,
        this._me === this.personDetails.id,
        MgtPersonCard.config
      );
    }

    this.loadSections();

    this._isStateLoading = false;
  }

  /**
   * Send a chat message to the user from the quick message input.
   *
   * @protected
   * @returns {void}
   * @memberof MgtPersonCard
   */
  protected sendQuickMessage = async (): Promise<void> => {
    const message = this._chatInput.trim();
    if (!message?.length) {
      return;
    }
    const person = this.personDetails as User;
    const user = this._me.userPrincipalName;
    this.isSending = true;

    const chat = await createChat(this._graph, person.userPrincipalName, user);

    const messageData = {
      body: {
        content: message
      }
    };
    await sendMessage(this._graph, chat.id, messageData);
    this.isSending = false;
    this.clearInputData();
  };

  /**
   * Use the mailto: protocol to initiate a new email to the user.
   *
   * @protected
   * @memberof MgtPersonCard
   */
  protected emailUser = () => {
    const user = this.internalPersonDetails;
    if (user) {
      const email = getEmailFromGraphEntity(user);
      if (email) {
        window.open('mailto:' + email, '_blank', 'noreferrer');
      }
    }
  };

  private get hasPhone(): boolean {
    const user = this.personDetails as User;
    const person = this.personDetails as microsoftgraph.Person;
    return Boolean(user?.businessPhones?.length) || Boolean(person?.phones?.length);
  }

  /**
   * Use the tel: protocol to initiate a new call to the user.
   *
   * @protected
   * @memberof MgtPersonCard
   */
  protected callUser = () => {
    const user = this.personDetails as User;
    const person = this.personDetails as microsoftgraph.Person;

    if (user?.businessPhones?.length) {
      const phone = user.businessPhones[0];
      if (phone) {
        window.open('tel:' + phone, '_blank', 'noreferrer');
      }
    } else if (person?.phones?.length) {
      const businessPhones = this.getPersonBusinessPhones(person);
      const phone = businessPhones[0];
      if (phone) {
        window.open('tel:' + phone, '_blank', 'noreferrer');
      }
    }
  };

  /**
   * Initiate a chat message to the user via deeplink.
   *
   * @protected
   * @memberof MgtPersonCard
   */
  protected chatUser = (message: string = null) => {
    const user = this.personDetails as User;
    if (user?.userPrincipalName) {
      const users: string = user.userPrincipalName;

      let url = `https://teams.microsoft.com/l/chat/0/0?users=${users}`;
      if (message?.length) {
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
  };

  /**
   * Initiate a teams call with video with a user via deeplink.
   *
   * @protected
   * @memberof MgtPersonCard
   */
  protected videoCallUser = () => {
    const user = this.personDetails as User;
    if (user?.userPrincipalName) {
      const users: string = user.userPrincipalName;

      const url = `https://teams.microsoft.com/l/call/0/0?users=${users}&withVideo=true`;

      const openWindow = () => window.open(url, '_blank');

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
  };

  /**
   * Display the expanded details panel.
   *
   * @protected
   * @memberof MgtPersonCard
   */
  protected showExpandedDetails = () => {
    const root = this.renderRoot.querySelector('.root');
    if (root?.animate) {
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
  };

  private loadSections() {
    this.sections = [];

    if (!this.internalPersonDetails) {
      return;
    }

    const contactSections = new MgtContact(this.internalPersonDetails as User);
    if (contactSections.hasData) {
      this.sections.push(contactSections);
    }

    if (!this._cardState) {
      return;
    }

    const { person, directReports, messages, files, profile } = this._cardState;

    if (MgtPersonCard.config.sections.organization && (person?.manager || directReports?.length)) {
      this.sections.push(new MgtOrganization(this._cardState, this._me));
    }

    if (MgtPersonCard.config.sections.mailMessages && messages?.length) {
      this.sections.push(new MgtMessages(messages));
    }

    if (MgtPersonCard.config.sections.files && files?.length) {
      this.sections.push(new MgtFileList());
    }

    if (MgtPersonCard.config.sections.profile && profile) {
      const profileSection = new MgtProfile(profile);
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
    return person?.personImage ? person.personImage : null;
  }

  private clearInputData() {
    this._chatInput = '';
    this.requestUpdate();
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

  private updateCurrentSection(section: CardSection) {
    if (section) {
      const sectionName = section.tagName.toLowerCase();
      const tabs: HTMLElement = this.renderRoot.querySelector(`#${sectionName}-Tab`);
      tabs.click();
    }
    const panels = this.renderRoot.querySelectorAll('fluent-tab-panel');
    for (const target of panels) {
      target.scrollTop = 0;
    }
    this._currentSection = section;
    this.requestUpdate();
  }

  private handleSectionScroll(e: WheelEvent) {
    const panels = this.renderRoot.querySelectorAll('fluent-tab-panel');
    for (const target of panels) {
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

  private readonly sendQuickMessageOnEnter = (e: KeyboardEvent) => {
    if (e.code === 'Enter') {
      void this.sendQuickMessage();
    }
  };

  private readonly handleGoBack = (e: KeyboardEvent) => {
    if (e.code === 'Enter') {
      void this.goBack();
    }
  };
}
