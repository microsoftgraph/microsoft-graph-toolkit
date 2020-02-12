/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property, TemplateResult } from 'lit-element';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import { getEmailFromGraphEntity } from '../../utils/GraphHelpers';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';
import { MgtPerson } from '../mgt-person/mgt-person';
import { MgtTemplatedComponent } from '../templatedComponent';
import { styles } from './mgt-person-card-css';
/**
 * Web Component used to show detailed data for a person in the
 * Microsoft Graph
 *
 * @export
 * @class MgtPersonCard
 * @extends {MgtTemplatedComponent}
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
   * @type {(MicrosoftGraph.User
   *     | MicrosoftGraph.Person
   *     | MicrosoftGraph.Contact)}
   * @memberof MgtPersonCard
   */
  @property({
    attribute: 'person-details',
    type: Object
  })
  public personDetails: MicrosoftGraph.User | MicrosoftGraph.Person | MicrosoftGraph.Contact;

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
  public attributeChangedCallback(name, oldValue, newValue) {
    super.attributeChangedCallback(name, oldValue, newValue);

    if (name === 'is-expanded' && oldValue !== newValue) {
      this.isExpanded = false;
    }
  }

  /**
   * Invoked when the element is first updated. Implement to perform one time
   * work on the element after update.
   *
   * Setting properties inside this method will trigger the element to update
   * again after this update cycle completes.
   *
   * * @param _changedProperties Map of changed properties with old values
   */
  public firstUpdated() {
    Providers.onProviderUpdated(() => this.loadData());
    this.loadData();
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    if (this.personDetails) {
      const user = this.personDetails;

      let department: TemplateResult;
      let jobTitle: TemplateResult;

      if (user.department) {
        department = html`
          <div class="department">${user.department}</div>
        `;
      }

      if (user.jobTitle) {
        jobTitle = html`
          <div class="job-title">${user.jobTitle}</div>
        `;
      }

      const image = this.getImage();

      return html`
        <div class="root" @click=${this.handleClose}>
          <div class="default-view">
            ${this.renderTemplate('default', { person: this.personDetails, personImage: image }) ||
              html`
                <mgt-person
                  class="person-image"
                  .personDetails=${this.personDetails}
                  .personImage=${image}
                ></mgt-person>
                <div class="details">
                  <div class="display-name">${user.displayName}</div>
                  ${jobTitle} ${department}
                  <div class="base-icons">
                    ${this.renderIcons()}
                  </div>
                </div>
              `}
          </div>
          <div class="additional-details-container">
            ${this.renderAdditionalDetails()}
          </div>
        </div>
      `;
    }
  }

  private renderIcons() {
    if (this.isExpanded === true) {
      return html``;
    } else {
      const user = this.personDetails;
      let chat: TemplateResult;
      let email: TemplateResult;
      let phone: TemplateResult;

      if ((user as MicrosoftGraph.User).mailNickname) {
        chat = html`
          <div class="icon" @click=${this._chatUser}>
            ${getSvg(SvgIcon.Chat, '#666666')}
          </div>
        `;
      }
      if (getEmailFromGraphEntity(user)) {
        email = html`
          <div class="icon" @click=${this._emailUser}>
            ${getSvg(SvgIcon.Email, '#666666')}
          </div>
        `;
      }
      if ((user as MicrosoftGraph.User).businessPhones && (user as MicrosoftGraph.User).businessPhones.length > 0) {
        phone = html`
          <div class="icon" @click=${this._callUser}>
            ${getSvg(SvgIcon.Phone, '#666666')}
          </div>
        `;
      }
      return html`
        ${chat} ${email} ${phone}
      `;
    }
  }

  private renderAdditionalDetails() {
    if (this.isExpanded === true) {
      const user = this.personDetails;

      let phone: TemplateResult;
      let email: TemplateResult;
      let location: TemplateResult;
      let chat: TemplateResult;

      if ((user as MicrosoftGraph.User).businessPhones && (user as MicrosoftGraph.User).businessPhones.length > 0) {
        phone = html`
          <div class="details-icon" @click=${this._callUser}>
            ${getSvg(SvgIcon.SmallPhone, '#666666')}
            <span class="link-subtitle data">${(user as MicrosoftGraph.User).businessPhones[0]}</span>
          </div>
        `;
      }

      if (getEmailFromGraphEntity(user)) {
        email = html`
          <div class="details-icon" @click=${this._emailUser}>
            ${getSvg(SvgIcon.SmallEmail, '#666666')}
            <span class="link-subtitle data">${getEmailFromGraphEntity(user)}</span>
          </div>
        `;
      }

      if ((user as MicrosoftGraph.User).mailNickname) {
        chat = html`
          <div class="details-icon" @click=${this._chatUser}>
            ${getSvg(SvgIcon.SmallChat, '#666666')}
            <span class="link-subtitle data">${(user as MicrosoftGraph.User).mailNickname}</span>
          </div>
        `;
      }

      if (user.officeLocation) {
        location = html`
          <div class="details-icon">
            ${getSvg(SvgIcon.SmallLocation, '#666666')}<span class="normal-subtitle data">${user.officeLocation}</span>
          </div>
        `;
      }
      const renderAdditionalSection: boolean = this.templates && this.templates['additional-details'];

      return html`
        <div class="additional-details-info">
          <div class="contact-text">Contact</div>
          <div class="additional-details-row">
            <div class="additional-details-item">
              <div class="icons">
                ${chat} ${email} ${phone} ${location}
              </div>
              ${renderAdditionalSection
                ? html`
                    <div class="section-divider"></div>
                    <div class="custom-section">
                      ${this.renderTemplate('additional-details', {
                        person: this.personDetails,
                        personImage: this.getImage()
                      })}
                    </div>
                  `
                : null}
            </div>
          </div>
        </div>
      `;
    } else {
      return html`
        <div class="additional-details-button" @click=${this._showAdditionalDetails}>
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
  }

  private _showAdditionalDetails(e: Event) {
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

  private _callUser(e: Event) {
    const user = this.personDetails;
    let phone: string;

    if ((user as MicrosoftGraph.User).businessPhones && (user as MicrosoftGraph.User).businessPhones.length > 0) {
      phone = (user as MicrosoftGraph.User).businessPhones[0];
    }
    e.stopPropagation();
    window.open('tel:' + phone, '_blank');
  }

  private _emailUser(e: Event) {
    const user = this.personDetails;
    let email;

    if (getEmailFromGraphEntity(user)) {
      email = getEmailFromGraphEntity(user);
    }
    e.stopPropagation();
    window.open('mailto:' + email, '_blank');
  }

  private _chatUser(e: Event) {
    const user = this.personDetails;
    let chat: string;

    if ((user as MicrosoftGraph.User).mailNickname) {
      chat = (user as MicrosoftGraph.User).mailNickname;
    }
    e.stopPropagation();
    window.open('sip:' + chat, '_blank');
  }

  private async loadData() {
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

  private handleClose(e: Event) {
    e.stopPropagation();
  }

  private getImage(): string {
    if (this.personImage && this.personImage !== '@') {
      return this.personImage;
    } else if (this.personDetails && (this.personDetails as any).personImage) {
      return (this.personDetails as any).personImage;
    }
    return null;
  }
}
