/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property, TemplateResult } from 'lit-element';
import { findPerson, findUserByEmail, getEmailFromGraphEntity } from '../../graph/graph.people';
import { getContactPhoto, getUserPhoto, myPhoto } from '../../graph/graph.photos';
import { getUser, getUserWithPhoto } from '../../graph/graph.user';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';
import { IDynamicPerson, MgtPerson } from '../mgt-person/mgt-person';
import { MgtTemplatedComponent } from '../templatedComponent';
import { styles } from './mgt-person-card-css';

/**
 * Web Component used to show detailed data for a person in the
 * Microsoft Graph
 *
 * @export
 * @class MgtPersonCard
 * @extends {MgtTemplatedComponent}
 *
 * @cssprop --font-size - {Length} Font size
 * @cssprop --font-weight - {Length} Font weight
 * @cssprop --height - {String} Height
 * @cssprop --background-color - {Color} Background color
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
   * Set the person details to render
   *
   * @type {IDynamicPerson}
   * @memberof MgtPersonCard
   */
  @property({
    attribute: 'person-details',
    type: Object
  })
  public personDetails: IDynamicPerson;

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
      <div class="display-name">${person.displayName}</div>
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
    const contactPerson = person as MicrosoftGraph.Contact;

    let chat: TemplateResult;
    let email: TemplateResult;
    let phone: TemplateResult;

    if (userPerson.userPrincipalName || personPerson.userPrincipalName) {
      chat = html`
        <div class="icon" @click=${this.chatUser}>
          ${getSvg(SvgIcon.Chat, '#666666')}
        </div>
      `;
    }
    if (getEmailFromGraphEntity(person)) {
      email = html`
        <div class="icon" @click=${() => this.emailUser()}>
          ${getSvg(SvgIcon.Email, '#666666')}
        </div>
      `;
    }
    if (
      (userPerson.businessPhones && userPerson.businessPhones.length > 0) ||
      (personPerson.phones && personPerson.phones.length > 0) ||
      (contactPerson.businessPhones && contactPerson.businessPhones.length > 0)
    ) {
      phone = html`
        <div class="icon" @click=${this.callUser}>
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
  }

  /**
   * Render additional details for the person.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCard
   */
  protected renderAdditionalDetails(person?: IDynamicPerson): TemplateResult {
    if (!this.hasTemplate('additional-details')) {
      return null;
    }

    person = person || this.personDetails;
    const personImage = this.getImage();
    const additionalDetailsTemplate = this.renderTemplate('additional-details', { person, personImage });

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

    if (this.hasTemplate('contact-details')) {
      return this.renderTemplate('contact-details', { person });
    }

    const userPerson = person as MicrosoftGraph.User;

    let phone: TemplateResult;
    let email: TemplateResult;
    let location: TemplateResult;
    let chat: TemplateResult;

    if (userPerson.businessPhones && userPerson.businessPhones.length > 0) {
      phone = html`
        <div class="details-icon" @click=${this.callUser}>
          ${getSvg(SvgIcon.SmallPhone, '#666666')}
          <span class="link-subtitle data">${userPerson.businessPhones[0]}</span>
        </div>
      `;
    }

    if (getEmailFromGraphEntity(person)) {
      email = html`
        <div class="details-icon" @click=${() => this.emailUser()}>
          ${getSvg(SvgIcon.SmallEmail, '#666666')}
          <span class="link-subtitle data">${getEmailFromGraphEntity(person)}</span>
        </div>
      `;
    }

    if (userPerson.mailNickname) {
      chat = html`
        <div class="details-icon" @click=${this.chatUser}>
          ${getSvg(SvgIcon.SmallChat, '#666666')}
          <span class="link-subtitle data">${userPerson.mailNickname}</span>
        </div>
      `;
    }

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
    const provider = Providers.globalProvider;
    const graph = provider.graph.forComponent(this);

    // check if user is signed in
    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    // check if personDetail already populated
    if (this.personDetails) {
      if (this.personDetails.id) {
        const user = this.personDetails as MicrosoftGraph.User;
        const person = this.personDetails as MicrosoftGraph.Person;
        const contact = this.personDetails as MicrosoftGraph.Contact;

        // check if there's userPrincipalName in personDetails
        if (user.userPrincipalName || person.userPrincipalName) {
          // check if there's email present in personDetails
          if (user.mail || person.scoredEmailAddresses || contact.emailAddresses) {
            return;
          }
        } else {
          this.personDetails = await getUser(graph, this.personDetails.id);
        }
      }
    }

    if (this.inheritDetails) {
      let parent = this.parentElement;
      while (parent && parent.tagName !== 'MGT-PERSON') {
        parent = parent.parentElement;
      }

      if (parent && (parent as MgtPerson).personDetails) {
        this.personDetails = (parent as MgtPerson).personDetails;
        this.personImage = (parent as MgtPerson).personImage;
      }
    } else {
      // Use userId or 'me' query to get the person and image
      if (this.userId || this.personQuery === 'me') {
        const person = await getUserWithPhoto(graph, this.userId);

        this.personDetails = person;
        this.personImage = this.getImage();
        return;
      }

      // Use the personQuery to find our person.
      if (this.personQuery) {
        const people = await findPerson(graph, this.personQuery);

        if (people && people.length) {
          this.personDetails = people[0];

          if (this.personImage === '@') {
            this.loadImage();
          }
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
    if (user && user.businessPhones && user.businessPhones.length) {
      const phone = user.businessPhones[0];
      if (phone) {
        window.open('tel:' + phone, '_blank');
      }
    }
  }

  /**
   * Use the sip: protocol to initiate a chat message to the user.
   *
   * @protected
   * @memberof MgtPersonCard
   */
  protected chatUser() {
    const user = this.personDetails as MicrosoftGraph.User;
    if (user && user.mailNickname) {
      const chat = user.mailNickname;
      window.open('sip:' + chat, '_blank');
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

  private async loadImage() {
    const provider = Providers.globalProvider;
    const graph = provider.graph.forComponent(this);
    const person = this.personDetails;
    let image: string;

    if ((person as MicrosoftGraph.Person).userPrincipalName) {
      // try to find a user by userPrincipalName
      const userPrincipalName = (person as MicrosoftGraph.Person).userPrincipalName;
      image = await getUserPhoto(graph, userPrincipalName);
    } else {
      // try to find a user by e-mail
      const email = getEmailFromGraphEntity(person);
      if (email) {
        const users = await findUserByEmail(graph, email);
        if (users && users.length) {
          // Check for an OrganizationUser
          const orgUser = users.find(p => {
            return (p as any).personType && (p as any).personType.subclass === 'OrganizationUser';
          });
          if (orgUser) {
            // Lookup by userId
            const userId = (users[0] as MicrosoftGraph.Person).scoredEmailAddresses[0].address;
            image = await getUserPhoto(graph, userId);
          } else {
            // Lookup by contactId
            const contactId = users[0].id;
            image = await getContactPhoto(graph, contactId);
          }
        }
      }
    }

    if (image) {
      this.personDetails.personImage = image;
      this.personImage = image;
    }
  }

  private getImage(): string {
    if (this.personImage && this.personImage !== '@') {
      return this.personImage;
    }

    const person = this.personDetails;
    return person && person.personImage ? person.personImage : null;
  }
}
