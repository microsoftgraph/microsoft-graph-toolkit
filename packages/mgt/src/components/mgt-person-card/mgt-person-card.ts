/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property, TemplateResult } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { findPeople, getEmailFromGraphEntity } from '../../graph/graph.people';
import { getPersonImage } from '../../graph/graph.photos';
import { getUserWithPhoto } from '../../graph/graph.user';
import { IDynamicPerson } from '../../graph/types';
import { Providers, ProviderState, MgtTemplatedComponent } from '@microsoft/mgt-element';
import { TeamsHelper } from '../../utils/TeamsHelper';
import { MgtPerson } from '../mgt-person/mgt-person';
import { styles } from './mgt-person-card-css';
import { BasePersonCardSection } from './sections/BasePersonCardSection';
import { MgtPersonCardContact } from './sections/mgt-person-card-contact/mgt-person-card-contact';
// import { MgtPersonCardFiles } from './sections/mgt-person-card-files/mgt-person-card-files';
import { MgtPersonCardMessages } from './sections/mgt-person-card-messages/mgt-person-card-messages';
import { MgtPersonCardOrganization } from './sections/mgt-person-card-organization/mgt-person-card-organization';
import { MgtPersonCardProfile } from './sections/mgt-person-card-profile/mgt-person-card-profile';
import { Presence } from '@microsoft/microsoft-graph-types-beta';
import { getUserPresence } from '../../graph/graph.presence';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';

/**
 * Web Component used to show detailed data for a person in the Microsoft Graph
 *
 * @export
 * @class MgtPersonCard
 * @extends {MgtTemplatedComponent}
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
    this.personImage = null;
    this.sections.forEach(s => (s.personDetails = value));
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
   * user-id property allows developer to use id value for component
   * @type {string}
   */
  @property({
    attribute: 'user-id'
  })
  public userId: string;

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
   * @type {MicrosoftGraphBeta.Presence}
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

  private _history: IDynamicPerson[];
  private _chatInput: string;
  private _currentSection: BasePersonCardSection;
  private _personDetails: IDynamicPerson;

  constructor() {
    super();
    this._chatInput = '';
    this._currentSection = null;
    this._history = [];
    this.sections = [
      new MgtPersonCardContact(),
      new MgtPersonCardOrganization(),
      new MgtPersonCardMessages(),
      // new MgtPersonCardFiles(),
      new MgtPersonCardProfile()
    ];
  }

  /**
   * Invoked each time the custom element is appended into a document-connected element
   *
   * @memberof MgtPersonCard
   */
  public connectedCallback() {
    super.connectedCallback();
    this.addEventListener('click', e => e.stopPropagation());
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

    if (name === 'is-expanded') {
      this.isExpanded = false;
    }

    switch (name) {
      case 'person-query':
      case 'user-id':
        this.personDetails = null;
        this.requestStateUpdate();
        break;
    }
  }

  /**
   * Navigate the card to a different user.
   *
   * @protected
   * @memberof MgtPersonCard
   */
  public navigate(person: IDynamicPerson): void {
    this._history.push(this.personDetails);

    this.sections.forEach((s: BasePersonCardSection) => {
      s.clearState();
      s.requestUpdate();
    });

    this.personDetails = person;
    this._currentSection = null;
  }

  /**
   * Navigate the card back to the previous user
   *
   * @returns {void}
   * @memberof MgtPersonCard
   */
  public goBack(): void {
    if (!this._history || !this._history.length) {
      return;
    }
    this.personDetails = this._history.pop();
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    // Handle no data
    if (!this.personDetails) {
      return this.renderNoData();
    }

    const person = this.personDetails;
    const image = this.getImage();
    const presence = this.personPresence;
    const showPresence = this.showPresence;

    // Check for a default template.
    // tslint:disable-next-line: no-string-literal
    if (this.hasTemplate('default')) {
      return this.renderTemplate('default', {
        person: this.personDetails,
        personImage: image
      });
    }

    const navigationTemplate =
      this._history && this._history.length
        ? html`
            <div class="nav">
              <div class="nav__back" @click=${() => this.goBack()}>
                ${getSvg(SvgIcon.Back)}
              </div>
            </div>
          `
        : null;

    // Check for a person-details template
    let personDetailsTemplate = this.renderTemplate('person-details', {
      person: this.personDetails,
      personImage: image
    });
    if (!personDetailsTemplate) {
      const personImageTemplate = this.renderPersonImage(image, presence, showPresence);
      const personNameTemplate = this.renderPersonName(person);
      const personTitleTemplate = this.renderPersonTitle(person);
      const personSubtitleTemplate = this.renderPersonSubtitle(person);
      const contactIconsTemplate = this.renderContactIcons(person);

      personDetailsTemplate = html`
        <div class="image">
          ${personImageTemplate}
        </div>
        <div class="details">
          ${personNameTemplate} ${personTitleTemplate} ${personSubtitleTemplate}
        </div>
        <div class="base-icons">
          ${contactIconsTemplate}
        </div>
      `;
    }

    const expandedDetailsTemplate = this.isExpanded ? this.renderExpandedDetails() : this.renderExpandedDetailsButton();

    return html`
      <div class="root">
        ${navigationTemplate}
        <div class="person-details-container">
          ${personDetailsTemplate}
        </div>
        <div class="expanded-details-container">
          ${expandedDetailsTemplate}
        </div>
      </div>
    `;
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
  protected renderPersonImage(imageSrc?: string, presence?: Presence, showPresence?: boolean): TemplateResult {
    imageSrc = imageSrc || this.getImage();
    presence = presence || this.personPresence;
    showPresence = showPresence || this.showPresence;
    const avatarSize = 'large';
    return html`
      <mgt-person
        class="person-image"
        .personDetails=${this.personDetails}
        .personImage=${imageSrc}
        .personPresence=${presence}
        .showPresence=${showPresence}
        .avatarSize=${avatarSize}
      ></mgt-person>
    `;
  }

  /**
   * Render the display name and persona details (e.g. department, job title) for a person.
   *
   * @protected
   * @param {IDynamicPerson} [person]
   * @returns {TemplateResult}
   * @memberof MgtPersonCard
   */
  protected renderPersonName(person?: IDynamicPerson): TemplateResult {
    person = person || this.personDetails;
    return html`
      <div class="display-name" title="${person.displayName}">${person.displayName}</div>
    `;
  }

  /**
   * Render person title.
   *
   * @protected
   * @param {IDynamicPerson} person
   * @returns {TemplateResult}
   * @memberof MgtPersonCard
   */
  protected renderPersonTitle(person?: IDynamicPerson): TemplateResult {
    person = person || this.personDetails;
    if (!person.jobTitle) {
      return;
    }
    return html`
      <div class="job-title">${person.jobTitle}</div>
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
    person = person || this.personDetails;
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
      return html``;
    }

    person = person || this.personDetails;
    const userPerson = person as MicrosoftGraph.User;

    // Email
    let email: TemplateResult;
    if (getEmailFromGraphEntity(person)) {
      email = html`
        <div class="icon" @click=${() => this.emailUser()}>
          ${getSvg(SvgIcon.SmallEmail)}
          <span>Send email</span>
        </div>
      `;
    }

    // Chat
    let chat: TemplateResult;
    if (userPerson.userPrincipalName) {
      chat = html`
        <div class="icon" @click=${() => this.chatUser()}>
          ${getSvg(SvgIcon.SmallChat)}
          <span>Start chat</span>
        </div>
      `;
    }

    return html`
      ${email} ${chat}
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
      <div class="expanded-details-button" @click=${() => this.showExpandedDetails()}>
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
    person = person || this.personDetails;

    const sectionNavTemplate = this.renderSectionNavigation();
    const currentSectionTemplate = this.renderCurrentSection();

    return html`
      <div class="section-nav">
        ${sectionNavTemplate}
      </div>
      <div class="section-host">
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
    const currentSectionIndex = this._currentSection ? this.sections.indexOf(this._currentSection) : -1;

    const navIcons = this.sections.map((section, i, a) => {
      const classes = classMap({
        active: i === currentSectionIndex,
        'section-nav__icon': true
      });
      return html`
        <button class=${classes} @click=${() => this.updateCurrentSection(section)}>
          ${section.renderIcon()}
        </button>
      `;
    });

    const overviewClasses = classMap({
      active: currentSectionIndex === -1,
      'section-nav__icon': true
    });
    return html`
      <button class=${overviewClasses} @click=${() => this.updateCurrentSection(null)}>
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
    const compactTemplates = this.sections.map(
      section => html`
        <div class="section">
          <div class="section__header">
            <div class="section__title">${section.displayName}</div>
            <a class="section__show-more" @click=${() => this.updateCurrentSection(section)}>Show more</a>
          </div>
          <div class="section__content">${section.asCompactView()}</div>
        </div>
      `
    );

    return html`
      <div class="quick-message">
        <input
          type="text"
          class="quick-message__input"
          placeholder="Message ${this.personDetails.displayName}"
          .value=${this._chatInput}
          @input=${(e: Event) => {
            this._chatInput = (e.target as HTMLInputElement).value;
          }}
        />
        <button class="quick-message__send" @click=${() => this.sendQuickMessage()}>
          ${getSvg(SvgIcon.Send)}
        </button>
      </div>
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
    if (!this.personDetails && this.inheritDetails) {
      // User person details inherited from parent tree
      let parent = this.parentElement;
      while (parent && parent.tagName !== 'MGT-PERSON') {
        parent = parent.parentElement;
      }

      if (parent && (parent as MgtPerson).personDetails) {
        this.personDetails = (parent as MgtPerson).personDetails;
        this.personImage = (parent as MgtPerson).personImage;
      }
    }

    const provider = Providers.globalProvider;

    // check if user is signed in
    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    const graph = provider.graph.forComponent(this);

    // check if personDetail already populated
    if (this.personDetails) {
      const user = this.personDetails as MicrosoftGraph.User;
      const id = user.userPrincipalName || user.id;
      // if we have an id but no email, we should get data from the graph
      if (id && !getEmailFromGraphEntity(user)) {
        const person = await getUserWithPhoto(graph, id);
        this.personDetails = person;
        this.personImage = this.getImage();
      } else if (
        !this.personDetails.personImage &&
        ((this.fetchImage && !this.personImage) || this.personImage === '@')
      ) {
        // in some cases we might only have name or email, but need to find the image
        const image = await getPersonImage(graph, this.personDetails);
        if (image) {
          this.personDetails.personImage = image;
          this.personImage = image;
        }
      }
    } else if (this.userId || this.personQuery === 'me') {
      // Use userId or 'me' query to get the person and image
      const person = await getUserWithPhoto(graph, this.userId);

      this.personDetails = person;
      this.personImage = this.getImage();
    } else if (this.personQuery) {
      // Use the personQuery to find our person.
      const people = await findPeople(graph, this.personQuery);

      if (people && people.length) {
        this.personDetails = people[0];
        const image = await getPersonImage(graph, this.personDetails);
        if (image) {
          this.personDetails.personImage = image;
          this.personImage = image;
        }
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
          this.personPresence = await getUserPresence(graph, this.personDetails.id);
        } else {
          this.personPresence = defaultPresence;
        }
      } catch (_) {
        // set up a default Presence in case beta api changes or getting error code
        this.personPresence = defaultPresence;
      }
    }
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
    const user = this.personDetails;
    if (user) {
      const email = getEmailFromGraphEntity(user);
      if (email) {
        window.open('mailto:' + email, '_blank');
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
    const user = this.personDetails as MicrosoftGraph.User;
    const person = this.personDetails as microsoftgraph.Person;

    if (user && user.businessPhones && user.businessPhones.length) {
      const phone = user.businessPhones[0];
      if (phone) {
        window.open('tel:' + phone, '_blank');
      }
    } else if (person && person.phones && person.phones.length) {
      const businessPhones = this.getPersonBusinessPhones(person);
      const phone = businessPhones[0];
      if (phone) {
        window.open('tel:' + phone, '_blank');
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
    const user = this.personDetails as MicrosoftGraph.User;
    if (user && user.userPrincipalName) {
      const users: string = user.userPrincipalName;

      let url = `https://teams.microsoft.com/l/chat/0/0?users=${users}`;
      if (message && message.length) {
        url += `&message=${message}`;
      }

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
  }

  private getImage(): string {
    if (this.personImage && this.personImage !== '@') {
      return this.personImage;
    }

    const person = this.personDetails;
    return person && person.personImage ? person.personImage : null;
  }

  private getPersonBusinessPhones(person: MicrosoftGraph.Person): string[] {
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
}
