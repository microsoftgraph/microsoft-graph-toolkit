/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property, TemplateResult } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { MgtPerson } from '../../../components/mgt-person/mgt-person';
import { MgtTemplatedComponent } from '../../../components/templatedComponent';
import { findPeople, getEmailFromGraphEntity } from '../../../graph/graph.people';
import { getPersonImage } from '../../../graph/graph.photos';
import { getUserWithPhoto } from '../../../graph/graph.user';
import { IDynamicPerson } from '../../../graph/types';
import { Providers } from '../../../Providers';
import { ProviderState } from '../../../providers/IProvider';
import { getSvg, SvgIcon } from '../../../utils/SvgHelper';
import { TeamsHelper } from '../../../utils/TeamsHelper';
import { styles } from './mgt-person-card-css';
import { BasePersonCardSection } from './sections/BasePersonCardSection';
import { MgtPersonCardContact } from './sections/mgt-person-card-contact/mgt-person-card-contact';
import { MgtPersonCardEmails } from './sections/mgt-person-card-emails/mgt-person-card-emails';
import { MgtPersonCardFiles } from './sections/mgt-person-card-files/mgt-person-card-files';
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
@customElement('mgt-person-card-beta')
export class MgtPersonCardBeta extends MgtTemplatedComponent {
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
    this.sections.forEach(s => (s.personDetails = value));
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
   * foo
   *
   * @protected
   * @type {BasePersonCardSection[]}
   * @memberof MgtPersonCardBeta
   */
  protected sections: BasePersonCardSection[];

  private _personDetails: IDynamicPerson;
  private _currentSection: BasePersonCardSection;

  constructor() {
    super();
    this.sections = [
      new MgtPersonCardContact(),
      new MgtPersonCardOrganization(),
      new MgtPersonCardEmails(),
      new MgtPersonCardFiles(),
      new MgtPersonCardProfile()
    ];
    this._currentSection = null;
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
          <div class="base-icons">
            ${contactIconsTemplate}
          </div>
        </div>
      `;
    }

    const expandedDetailsTemplate = this.isExpanded ? this.renderExpandedDetails() : this.renderExpandedDetailsButton();

    return html`
      <div class="root">
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
      <mgt-person class="person-image" .personDetails=${this.personDetails} .personImage=${imageSrc}></mgt-person>
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
    const personPerson = person as MicrosoftGraph.Person;

    // Chat
    let chat: TemplateResult;
    if (userPerson.userPrincipalName) {
      chat = html`
        <div class="icon" @click=${() => this.chatUser()}>
          ${getSvg(SvgIcon.Chat, '#666666')}
        </div>
      `;
    }

    // Email
    let email: TemplateResult;
    if (getEmailFromGraphEntity(person)) {
      email = html`
        <div class="icon" @click=${() => this.emailUser()}>
          ${getSvg(SvgIcon.Email, '#666666')}
        </div>
      `;
    }

    // Phone
    let phone: TemplateResult;
    if (
      (userPerson.businessPhones && userPerson.businessPhones.length > 0) ||
      (personPerson.phones && personPerson.phones.length > 0)
    ) {
      phone = html`
        <div class="icon" @click=${() => this.callUser()}>
          ${getSvg(SvgIcon.Phone, '#666666')}
        </div>
      `;
    }

    return html`
      ${chat} ${email} ${phone}
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
        <svg
          class="expanded-details-svg"
          width="16"
          height="15"
          viewBox="0 0 16 15"
          fill="none"
          xmlns="http://www.w3.org/2000/svg"
        >
          <path d="M15 7L8.24324 13.7568L1.24324 6.75676" stroke="#3078CD" />
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
    /*
    const contactDetailsTemplate = this.renderContactDetails(person);
    const additionalDetailsTemplate = this.renderAdditionalDetails(person);
    const sectionDivider = additionalDetailsTemplate
      ? html`
          <div class="section-divider"></div>
        `
      : null;

    return html`
      ${contactDetailsTemplate} ${sectionDivider} ${additionalDetailsTemplate}
    `;
    */
  }

  /**
   * foo
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCardBeta
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
   * foo
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCardBeta
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
        <input type="text" class="quick-message__input" placeholder="Message ${this.personDetails.displayName}" />
        <button class="quick-message__send">
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
   * foo
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCardBeta
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
  protected chatUser() {
    const user = this.personDetails as MicrosoftGraph.User;
    if (user && user.userPrincipalName) {
      const users: string = user.userPrincipalName;
      const url = `https://teams.microsoft.com/l/chat/0/0?users=${users}`;
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
    this._currentSection = section;
    this.requestUpdate();
  }
}
