/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property, TemplateResult } from 'lit-element';
import { getEmailFromGraphEntity } from '../../graph/graph.people';
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
   * Gets or sets whether additional details section is rendered
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
   * Gets or sets whether additional details should be inherited from an mgt-person parent
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
   * Synchronizes property values when attributes change.
   *
   * @param {*} name
   * @param {*} oldValue
   * @param {*} newValue
   * @memberof MgtPersonCard
   */
  public attributeChangedCallback(name: string, oldValue: string, newValue: string) {
    super.attributeChangedCallback(name, oldValue, newValue);

    if (name === 'is-expanded' && oldValue !== newValue) {
      this.isExpanded = false;
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

    // Check for a details template
    let contentTemplate: TemplateResult = this.renderTemplate('details', {
      person: this.personDetails,
      personImage: image
    });
    if (!contentTemplate) {
      const personImageTemplate = this.renderPersonImage(image);
      const personDetailsTemplate = this.renderPersonDetails(person);

      contentTemplate = html`
        <div class="image">
          ${personImageTemplate}
        </div>
        <div class="details">
          ${personDetailsTemplate}
        </div>
      `;
    }

    const additionalDetailsTemplate = this.isExpanded
      ? this.renderAdditionalDetails()
      : this.renderAdditionalDetailsButton();

    return html`
      <div class="root" @click=${(e: Event) => this.handleClick(e)}>
        <div class="default-view">
          ${contentTemplate}
        </div>
        <div class="additional-details-container">
          ${additionalDetailsTemplate}
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
   * Render the person details (e.g. name, department, job title, icons)
   *
   * @protected
   * @param {IDynamicPerson} [person]
   * @returns {TemplateResult}
   * @memberof MgtPersonCard
   */
  protected renderPersonDetails(person?: IDynamicPerson): TemplateResult {
    person = person || this.personDetails;

    let department: TemplateResult;
    let jobTitle: TemplateResult;

    if (person.department) {
      department = html`
        <div class="department">${person.department}</div>
      `;
    }

    if (person.jobTitle) {
      jobTitle = html`
        <div class="job-title">${person.jobTitle}</div>
      `;
    }

    const contactIconsTemplate = this.renderContactIcons(person);

    return html`
      <div class="display-name">${person.displayName}</div>
      ${jobTitle} ${department}
      <div class="base-icons">
        ${contactIconsTemplate}
      </div>
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

    let chat: TemplateResult;
    let email: TemplateResult;
    let phone: TemplateResult;

    if (userPerson.mailNickname) {
      chat = html`
        <div class="icon" @click=${this.chatUser}>
          ${getSvg(SvgIcon.Chat, '#666666')}
        </div>
      `;
    }
    if (getEmailFromGraphEntity(person)) {
      email = html`
        <div class="icon" @click=${this.emailUser}>
          ${getSvg(SvgIcon.Email, '#666666')}
        </div>
      `;
    }
    if (userPerson.businessPhones && userPerson.businessPhones.length > 0) {
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
   * Render the button used to expand the additional details.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCard
   */
  protected renderAdditionalDetailsButton(): TemplateResult {
    return html`
      <div class="additional-details-button" @click=${(e: Event) => this.showAdditionalDetails(e)}>
        <svg
          class="additional-details-svg"
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
   * Render additional details for the person.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCard
   */
  protected renderAdditionalDetails(person?: IDynamicPerson): TemplateResult {
    person = person || this.personDetails;

    const contactDetailsTemplate = this.renderContactDetails();

    // Check for additional details template
    let additionalDetailsTemplate: TemplateResult = null;
    if (this.hasTemplate('additional-details')) {
      additionalDetailsTemplate = html`
        <div class="section-divider"></div>
        ${this.renderTemplate('additional-details', {
          person,
          personImage: this.getImage()
        })}
      `;
    }

    return html`
      <div class="additional-details-info">
        ${contactDetailsTemplate} ${additionalDetailsTemplate}
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
        <div class="details-icon" @click=${this.emailUser}>
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
      <div class="contact-text">Contact</div>
      <div class="additional-details-row">
        <div class="additional-details-item">
          <div class="icons">
            ${chat} ${email} ${phone} ${location}
          </div>
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
    if (this.inheritDetails) {
      let parent = this.parentElement;
      while (parent && parent.tagName !== 'MGT-PERSON') {
        parent = parent.parentElement;
      }

      if (parent && (parent as MgtPerson).personDetails) {
        this.personDetails = (parent as MgtPerson).personDetails;
        this.personImage = (parent as MgtPerson).personImage;
      }
    }

    if (this.personDetails) {
      return;
    }

    const provider = Providers.globalProvider;
    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }
  }

  private showAdditionalDetails(e: Event) {
    // handles clicking when parent is also activated by click
    e.stopPropagation();
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

  private callUser(e: Event) {
    const user = this.personDetails;
    let phone: string;

    if ((user as MicrosoftGraph.User).businessPhones && (user as MicrosoftGraph.User).businessPhones.length > 0) {
      phone = (user as MicrosoftGraph.User).businessPhones[0];
    }
    e.stopPropagation();
    window.open('tel:' + phone, '_blank');
  }

  private emailUser(e: Event) {
    const user = this.personDetails;
    let email;

    if (getEmailFromGraphEntity(user)) {
      email = getEmailFromGraphEntity(user);
    }
    e.stopPropagation();
    window.open('mailto:' + email, '_blank');
  }

  private chatUser(e: Event) {
    const user = this.personDetails;
    let chat: string;

    if ((user as MicrosoftGraph.User).mailNickname) {
      chat = (user as MicrosoftGraph.User).mailNickname;
    }
    e.stopPropagation();
    window.open('sip:' + chat, '_blank');
  }

  private handleClick(e: Event) {
    e.stopPropagation();
  }

  private getImage(): string {
    if (this.personImage && this.personImage !== '@') {
      return this.personImage;
    }

    const person = this.personDetails;
    return person && person.personImage ? person.personImage : null;
  }
}
