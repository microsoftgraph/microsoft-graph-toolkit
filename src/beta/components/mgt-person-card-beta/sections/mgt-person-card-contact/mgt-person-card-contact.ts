/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

// import * as GraphTypes from '@microsoft/microsoft-graph-types';
import * as GraphTypes from '@microsoft/microsoft-graph-types-beta';
import { customElement, html, TemplateResult } from 'lit-element';
import { getEmailFromGraphEntity } from '../../../../../graph/graph.people';
import { BasePersonCardSection } from '../BasePersonCardSection';
import { styles } from './mgt-person-card-contact-css';

/**
 * foo
 *
 * @interface IContectPart
 */
interface IContactPart {
  // tslint:disable-next-line: completed-docs
  title: string;
  // tslint:disable-next-line: completed-docs
  value: string;
  // tslint:disable-next-line: completed-docs
  onClick?: (e: Event) => void;
}

/**
 * foo
 *
 * @export
 * @class MgtPersonCardProfile
 * @extends {MgtTemplatedComponent}
 */
@customElement('mgt-person-card-contact')
export class MgtPersonCardContact extends BasePersonCardSection {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * foo
   *
   * @readonly
   * @type {string}
   * @memberof MgtPersonCardContact
   */
  public get displayName(): string {
    return 'Contact';
  }

  private _contactDetails: IContactPart[];

  /**
   * foo
   *
   * @returns {TemplateResult}
   * @memberof MgtPersonCardContact
   */
  public renderCompactView(): TemplateResult {
    return html`
      compact
    `;
  }

  /**
   * foo
   *
   * @returns {TemplateResult}
   * @memberof MgtPersonCardContact
   */
  public renderIcon(): TemplateResult {
    return html`
      icon
    `;
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    const templateParts = this._contactDetails ? this._contactDetails.map(part => this.renderContactPart(part)) : [];
    return html`
      <div class="title">About</div>
      ${templateParts}
    `;
  }

  /**
   * foo
   *
   * @protected
   * @param {IContactPart} part
   * @returns {TemplateResult}
   * @memberof MgtPersonCardContact
   */
  protected renderContactPart(part: IContactPart): TemplateResult {
    const valueTemplate = this.isLoadingState
      ? html`
          <div class="shimmer"></div>
        `
      : part.onClick
      ? html`
          <a @click=${(e: Event) => part.onClick(e)}>${part.value}</a>
        `
      : html`
          ${part.value}
        `;

    return html`
      <div class="part">
        <div class="part__title">${part.title}</div>
        <div class="part__value">${valueTemplate}</div>
      </div>
    `;
  }

  /**
   * foo
   *
   * @protected
   * @returns {IContactPart[]}
   * @memberof MgtPersonCardContact
   */
  protected async loadState(): Promise<void> {
    if (!this.personDetails) {
      return;
    }

    const userPerson = this.personDetails as GraphTypes.User;
    const personPerson = this.personDetails as GraphTypes.Person;

    const contactParts: IContactPart[] = [];

    const email = getEmailFromGraphEntity(this.personDetails);
    if (email) {
      contactParts.push({
        onClick: (e: Event) => this.sendEmail(email),
        title: 'Email',
        value: email
      });
    }

    const chat = personPerson.userPrincipalName;
    if (chat) {
      contactParts.push({
        onClick: (e: Event) => this.sendChat(chat),
        title: 'Teams',
        value: chat
      });
    }

    const cellPhone = userPerson.mobilePhone;
    if (cellPhone) {
      contactParts.push({
        onClick: (e: Event) => this.sendCall(cellPhone),
        title: 'Cell Phone',
        value: cellPhone
      });
    }

    const department = this.personDetails.department;
    if (department) {
      contactParts.push({
        title: 'Department',
        value: department
      });
    }

    const title = this.personDetails.jobTitle;
    if (title) {
      contactParts.push({
        title: 'Title',
        value: title
      });
    }

    const officeLocation = this.personDetails.officeLocation;
    if (officeLocation) {
      contactParts.push({
        onClick: (e: Event) => this.showOfficeLocation(officeLocation),
        title: 'Office Location',
        value: officeLocation
      });
    }

    this._contactDetails = contactParts;
  }

  /**
   * foo
   *
   * @protected
   * @memberof MgtPersonCardContact
   */
  protected sendChat(chat: string): void {
    // foo
  }

  /**
   * foo
   *
   * @protected
   * @memberof MgtPersonCardContact
   */
  protected sendEmail(email: string): void {
    // foo
  }

  /**
   * foo
   *
   * @protected
   * @memberof MgtPersonCardContact
   */
  protected sendCall(phone: string): void {
    // foo
  }

  /**
   * foo
   *
   * @protected
   * @memberof MgtPersonCardContact
   */
  protected showOfficeLocation(officeLocation: string): void {
    // foo
  }
}
