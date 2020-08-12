/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property, TemplateResult } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { MgtPerson } from '../../components/mgt-person/mgt-person';
import { MgtTemplatedComponent } from '../../components/templatedComponent';
import { findPeople, getEmailFromGraphEntity } from '../../graph/graph.people';
import { getPersonImage } from '../../graph/graph.photos';
import { getMe, getUserWithPhoto } from '../../graph/graph.user';
import { IDynamicPerson } from '../../graph/types';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';
import { TeamsHelper } from '../../utils/TeamsHelper';
import { styles } from './mgt-person-card-css';
import { BasePersonCardSection } from './sections/BasePersonCardSection';
import { MgtPersonCardContact } from './sections/mgt-person-card-contact/mgt-person-card-contact';
import { MgtPersonCardFiles } from './sections/mgt-person-card-files/mgt-person-card-files';
import { MgtPersonCardMessages } from './sections/mgt-person-card-messages/mgt-person-card-messages';
import { MgtPersonCardOrganization } from './sections/mgt-person-card-organization/mgt-person-card-organization';
import './sections/mgt-person-card-profile/mgt-person-card-profile';
import { MgtPersonCardProfile } from './sections/mgt-person-card-profile/mgt-person-card-profile';

/**
 * Web Component used to show detailed data for a person in the
 * Microsoft Graph
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
      new MgtPersonCardFiles(),
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
                <svg width="16" height="16" viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg">
                  <path
                    d="M16 8.5H1.95312L8.10156 14.6484L7.39844 15.3516L0.046875 8L7.39844 0.648438L8.10156 1.35156L1.95312 7.5H16V8.5Z"
                  />
                </svg>
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
      const personImageTemplate = this.renderPersonImage(image);
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
        <div></div>
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
  protected renderPersonImage(imageSrc?: string): TemplateResult {
    imageSrc = imageSrc || this.getImage();
    return html`
      <mgt-person
        class="person-image"
        .personDetails=${this.personDetails}
        .personImage=${imageSrc}
        .showPresence=${true}
        avatar-size="large"
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
          <svg width="14" height="10" viewBox="0 0 14 10" xmlns="http://www.w3.org/2000/svg">
            <path
              fill-rule="evenodd"
              clip-rule="evenodd"
              d="M11.8473 1H2.04886L6.47364 4.18916C6.64273 4.31103 6.86969 4.31522 7.04316 4.19968L11.8473 1ZM1 1.47671V9H13V1.43376L7.59749 5.03198C7.07706 5.3786 6.39621 5.36601 5.88894 5.0004L1 1.47671ZM0 1C0 0.447715 0.447715 0 1 0H13C13.5523 0 14 0.447715 14 1V9C14 9.55228 13.5523 10 13 10H1C0.447716 10 0 9.55229 0 9V1Z"
            />
          </svg>
          <span>Send email</span>
        </div>
      `;
    }

    // Chat
    let chat: TemplateResult;
    if (userPerson.userPrincipalName) {
      chat = html`
        <div class="icon" @click=${() => this.chatUser()}>
          <svg width="13" height="13" viewBox="0 0 13 13" xmlns="http://www.w3.org/2000/svg">
            <path
              fill-rule="evenodd"
              clip-rule="evenodd"
              d="M9.5781 8.26403L9.57811 8.26403C9.68611 8.29867 9.8455 8.36364 10.0125 8.50479C10.1405 8.61294 10.2409 8.73883 10.3235 8.86465L11.9634 10.9944C11.9701 10.9924 11.9753 10.9904 11.9785 10.9889C11.9841 10.9864 11.9918 10.9823 12 10.9768V1.32284C12 1.18078 11.8731 1 11.6207 1H1.37926C1.12692 1 1 1.18078 1 1.32284V7.37377C1 7.45357 1.01415 7.49036 1.02102 7.50507C1.02778 7.51955 1.04342 7.54689 1.09159 7.58705L1.10485 7.5981L1.11771 7.6096C1.13526 7.62529 1.21707 7.69076 1.33937 7.76027C1.46122 7.82952 1.58119 7.87864 1.67944 7.89966L1.69102 7.90214L1.691 7.9022C3.32106 8.27116 6.2626 8.27688 8.67896 8.18036L8.69908 8.17955L8.71921 8.17956L8.7627 8.17954C9.01362 8.17932 9.31313 8.17907 9.5781 8.26403ZM11.2376 11.6908L9.50506 9.44081C9.39493 9.26422 9.32445 9.23285 9.27276 9.21627C9.17534 9.18504 9.0401 9.17966 8.71888 9.17956C6.31879 9.27543 3.24831 9.27999 1.47024 8.87753C1.01314 8.77974 0.600449 8.48852 0.451238 8.35513C0.15593 8.10893 0 7.78958 0 7.37377V1.32284C0 0.593147 0.610664 0 1.37926 0H11.6207C12.3893 0 13 0.593147 13 1.32284V11.0441C13 11.4686 12.6828 11.7689 12.3881 11.9012C12.1048 12.0284 11.673 12.0728 11.3387 11.7993C11.3003 11.7679 11.2678 11.7301 11.2376 11.6908ZM3 3.5C3 3.22386 3.22386 3 3.5 3H9.5C9.77614 3 10 3.22386 10 3.5C10 3.77614 9.77614 4 9.5 4H3.5C3.22386 4 3 3.77614 3 3.5ZM3.5 5C3.22386 5 3 5.22386 3 5.5C3 5.77614 3.22386 6 3.5 6H6.5C6.77614 6 7 5.77614 7 5.5C7 5.22386 6.77614 5 6.5 5H3.5Z"
            />
          </svg>
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
        <svg width="15" height="8" viewBox="0 0 15 8" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M14 1L7.5 7L1 1" stroke="#3078CD" />
        </svg>
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
        <svg xmlns="http://www.w3.org/2000/svg">
          <path
            d="M12 4V9H2V4H12ZM11 8V5H3V8H11ZM13 4H18V9H13V4ZM17 8V5H14V8H17ZM8 15V10H18V15H8ZM9 11V14H17V11H9ZM2 15V10H7V15H2ZM3 11V14H6V11H3Z"
          />
        </svg>
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
          <svg xmlns="http://www.w3.org/2000/svg">
            <path
              d="M4.27144 8.99999L1.72572 2.45387C1.54854 1.99826 1.9928 1.56256 2.43227 1.71743L2.50153 1.74688L16.0015 8.49688C16.3902 8.69122 16.4145 9.22336 16.0744 9.45992L16.0015 9.50311L2.50153 16.2531C2.0643 16.4717 1.58932 16.0697 1.70282 15.6178L1.72572 15.5461L4.27144 8.99999L1.72572 2.45387L4.27144 8.99999ZM3.3028 3.4053L5.25954 8.43705L10.2302 8.43749C10.515 8.43749 10.7503 8.64911 10.7876 8.92367L10.7927 8.99999C10.7927 9.28476 10.5811 9.52011 10.3065 9.55736L10.2302 9.56249L5.25954 9.56205L3.3028 14.5947L14.4922 8.99999L3.3028 3.4053Z"
            />
          </svg>
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
   * Render additional details for the person.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCard
   */
  protected renderAdditionalDetails(person?: IDynamicPerson): TemplateResult {
    person = person || this.personDetails;
    const personImage = this.getImage();

    const additionalDetailsTemplate = this.hasTemplate('additional-details')
      ? this.renderTemplate('additional-details', { person, personImage })
      : html`
          <mgt-person-card-profile user-id="${person.id}"></mgt-person-card-profile>
        `;

    return html`
      <div class="expanded-details-info">
        ${additionalDetailsTemplate}
      </div>
    `;
  }

  /**
   * Render the contact info part of the additional details.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCard
   */
  protected renderContactDetails(person?: IDynamicPerson): TemplateResult {
    person = person || this.personDetails;
    const userPerson = person as MicrosoftGraph.User;
    const personPerson = person as MicrosoftGraph.Person;

    if (this.hasTemplate('contact-details')) {
      return this.renderTemplate('contact-details', { userPerson });
    }

    // Chat
    let chat: TemplateResult;
    if (userPerson.userPrincipalName) {
      chat = html`
        <div class="details-icon" @click=${() => this.chatUser()}>
          ${getSvg(SvgIcon.SmallChat, '#666666')}
          <span class="link-subtitle data">${userPerson.userPrincipalName}</span>
        </div>
      `;
    }

    // Email
    let email: TemplateResult;
    if (getEmailFromGraphEntity(person)) {
      email = html`
        <div class="details-icon" @click=${() => this.emailUser()}>
          ${getSvg(SvgIcon.SmallEmail, '#666666')}
          <span class="link-subtitle data">${getEmailFromGraphEntity(person)}</span>
        </div>
      `;
    }

    // Phone
    let phone: TemplateResult;
    if (userPerson.businessPhones && userPerson.businessPhones.length > 0) {
      phone = html`
        <div class="details-icon" @click=${() => this.callUser()}>
          ${getSvg(SvgIcon.SmallPhone, '#666666')}
          <span class="link-subtitle data">${userPerson.businessPhones[0]}</span>
        </div>
      `;
    } else if (personPerson.phones && personPerson.phones.length > 0) {
      const businessPhones = this.getPersonBusinessPhones(personPerson);
      phone = html`
        <div class="details-icon" @click=${() => this.callUser()}>
          ${getSvg(SvgIcon.SmallPhone, '#666666')}
          <span class="link-subtitle data">${businessPhones[0]}</span>
        </div>
      `;
    }

    // Location
    let location: TemplateResult;
    if (person.officeLocation) {
      location = html`
        <div class="details-icon">
          ${getSvg(SvgIcon.SmallLocation, '#666666')}<span class="normal-subtitle data">${person.officeLocation}</span>
        </div>
      `;
    }

    return html`
      <div class="expanded-details-info">
        <div class="contact-text">Contact</div>
        <div class="icons">
          ${chat} ${email} ${phone} ${location}
        </div>
      </div>
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
      } else if (this.personImage === '@' && !this.personDetails.personImage) {
        // in some cases we might only have name or email, but need to find the image
        // use @ for the image value to search for an image
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
