import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property } from 'lit-element';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import { getEmailFromGraphEntity } from '../../utils/graphHelpers';
import * as svgHelper from '../../utils/svgHelper';
import { PersonCardInteraction } from '../mgt-person/mgt-person';
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
   * @type {string}
   * @memberof MgtPersonCard
   */
  @property({
    attribute: 'is-extended',
    type: Boolean
  })
  public isExtended: boolean = false;

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

    if (name === 'is-extended' && oldValue !== newValue) {
      this.isExtended = false;
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

      // tslint:disable-next-line: one-variable-per-declaration
      let department, jobTitle;

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
      return html`
        <div class="root" @click=${this.handleClose}>
          <div class="default-view">
            ${this.renderTemplate('default', { person: this.personDetails }) ||
              html`
                <mgt-person .personDetails=${this.personDetails} .personImage=${this.personImage}></mgt-person>
                <div class="details">
                  <div class="display-name">${user.displayName}</div>
                  ${jobTitle} ${department}
                  <div class="base-icons">
                    ${this.renderIcons()}
                  </div>
                </div>
              `}
          </div>
          <div class="additional-details-container" @click=${this._showAdditionalDetails}>
            ${this.renderAdditionalDetails()}
          </div>
        </div>
      `;
    }
  }

  private renderIcons() {
    const user = this.personDetails;
    // tslint:disable-next-line: one-variable-per-declaration
    let chat, email, phone;
    if ((user as MicrosoftGraph.User).mailNickname) {
      chat = html`
        <div @mouseout=${this._unsetMouseOverState} @mouseover=${this._setMouseOverState} @click=${this._chatUser}>
          ${svgHelper.getSVG('chat', '#666666')}
        </div>
      `;
    }
    if (getEmailFromGraphEntity(user)) {
      email = html`
        <div @mouseout=${this._unsetMouseOverState} @mouseover=${this._setMouseOverState} @click=${this._emailUser}>
          ${svgHelper.getSVG('email', '#666666')}
        </div>
      `;
    }

    if ((user as MicrosoftGraph.User).businessPhones && (user as MicrosoftGraph.User).businessPhones.length > 0) {
      phone = html`
        <div @mouseout=${this._unsetMouseOverState} @mouseover=${this._setMouseOverState} @click=${this._callUser}>
          ${svgHelper.getSVG('phone', '#666666')}
        </div>
      `;
    }

    if (this.isExtended === true) {
      return html``;
    } else {
      return html`
        ${chat} ${email} ${phone}
      `;
    }
  }

  private renderAdditionalDetails() {
    const user = this.personDetails;

    // tslint:disable-next-line: one-variable-per-declaration
    let phone, email, location, chat;

    if ((user as MicrosoftGraph.User).businessPhones && (user as MicrosoftGraph.User).businessPhones.length > 0) {
      phone = html`
        <div class="details-icon" @click=${this._callUser}>
          ${svgHelper.getSVG('phone-small', '#666666')}
          <span class="link-subtitle data" @mouseout=${this._unsetMouseOverState} @mouseover=${this._setMouseOverState}
            >${(user as MicrosoftGraph.User).businessPhones[0]}</span
          >
        </div>
      `;
    }

    if (getEmailFromGraphEntity(user)) {
      email = html`
        <div class="details-icon" @click=${this._emailUser}>
          ${svgHelper.getSVG('email-small', '#666666')}
          <span class="link-subtitle data" @mouseout=${this._unsetMouseOverState} @mouseover=${this._setMouseOverState}
            >${getEmailFromGraphEntity(user)}</span
          >
        </div>
      `;
    }

    if ((user as MicrosoftGraph.User).mailNickname) {
      chat = html`
        <div class="details-icon" @click=${this._chatUser}>
          ${svgHelper.getSVG('chat-small', '#666666')}
          <span class="link-subtitle data" @mouseout=${this._unsetMouseOverState} @mouseover=${this._setMouseOverState}
            >${(user as MicrosoftGraph.User).mailNickname}</span
          >
        </div>
      `;
    }

    if (user.officeLocation) {
      location = html`
        <div class="details-icon">
          ${svgHelper.getSVG('location-small', '#666666')}<span class="normal-subtitle data"
            >${user.officeLocation}</span
          >
        </div>
      `;
    }

    if (this.isExtended === true) {
      return html`
        <div class="additional-details-info">
          <div class="contact-text">Contact</div>
          <div class="additional-details-row">
            <div class="additional-details-item">
              <div class="icons">
                ${chat} ${email} ${phone} ${location}
              </div>
              <div class="section-divider"></div>
              <div class="custom-section">
                ${this.renderTemplate('additional-details', null)}
              </div>
            </div>
          </div>
        </div>
      `;
    } else {
      return html`
        <div class="additional-details-button">
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

      this.isExtended = true;
    }
  }

  private _callUser(e: Event) {
    const user = this.personDetails;
    // tslint:disable-next-line: one-variable-per-declaration
    let phone;

    if ((user as MicrosoftGraph.User).businessPhones && (user as MicrosoftGraph.User).businessPhones.length > 0) {
      phone = (user as MicrosoftGraph.User).businessPhones[0];
    }
    e.stopPropagation();
    window.location.assign('tel:' + phone);
  }

  private _emailUser(e: Event) {
    const user = this.personDetails;
    let email;

    if (getEmailFromGraphEntity(user)) {
      email = getEmailFromGraphEntity(user);
    }
    e.stopPropagation();
    window.location.assign('mailto:' + email);
  }

  private _chatUser(e: Event) {
    const user = this.personDetails;
    // tslint:disable-next-line: one-variable-per-declaration
    let chat;

    if ((user as MicrosoftGraph.User).mailNickname) {
      chat = (user as MicrosoftGraph.User).mailNickname;
    }
    e.stopPropagation();
    window.location.assign('sip:' + chat);
  }

  private async loadData() {
    if (this.personDetails) {
      return;
    }

    const provider = Providers.globalProvider;

    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }
  }

  private _setMouseOverState(el) {
    el.target.classList.add('hover-state');
  }
  private _unsetMouseOverState(el) {
    el.target.classList.remove('hover-state');
  }

  private handleClose(e: Event) {
    e.stopPropagation();
  }
}
