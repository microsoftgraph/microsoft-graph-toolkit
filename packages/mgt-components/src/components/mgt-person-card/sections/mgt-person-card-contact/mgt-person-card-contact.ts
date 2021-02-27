/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { User } from '@microsoft/microsoft-graph-types';
import { customElement, html, TemplateResult } from 'lit-element';
import { TeamsHelper } from '@microsoft/mgt-element';
import { classMap } from 'lit-html/directives/class-map';

import { getEmailFromGraphEntity } from '../../../../graph/graph.people';
import { BasePersonCardSection } from '../BasePersonCardSection';
import { styles } from './mgt-person-card-contact-css';
import { getSvg, SvgIcon } from '../../../../utils/SvgHelper';
import { strings } from './strings';

/**
 * Represents a contact part and its metadata
 *
 * @interface IContectPart
 */
interface IContactPart {
  // tslint:disable-next-line: completed-docs
  icon: TemplateResult;
  // tslint:disable-next-line: completed-docs
  title: string;
  // tslint:disable-next-line: completed-docs
  value?: string;
  // tslint:disable-next-line: completed-docs
  onClick?: (e: Event) => void;
  // tslint:disable-next-line: completed-docs
  showCompact: boolean;
}

/**
 * The contact details subsection of the person card
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

  protected get strings() {
    return strings;
  }

  /**
   * Returns true if the component has data it can render
   *
   * @readonly
   * @abstract
   * @type {boolean}
   * @memberof BasePersonCardSection
   */
  public get hasData(): boolean {
    if (!this._contactParts) {
      return false;
    }

    const availableParts: IContactPart[] = Object.values(this._contactParts).filter((p: IContactPart) => !!p.value);

    return !!availableParts.length;
  }

  private _person: User;

  // tslint:disable: object-literal-sort-keys
  private _contactParts = {
    email: {
      icon: getSvg(SvgIcon.Email, '#929292'),
      onClick: () => this.sendEmail(),
      showCompact: true,
      title: 'Email'
    } as IContactPart,
    chat: {
      icon: getSvg(SvgIcon.Chat, '#929292'),
      onClick: () => this.sendChat(),
      showCompact: false,
      title: 'Teams'
    } as IContactPart,
    businessPhone: {
      icon: getSvg(SvgIcon.CellPhone, '#929292'),
      onClick: () => this.sendCall(),
      showCompact: true,
      title: 'Business Phone'
    } as IContactPart,
    cellPhone: {
      icon: getSvg(SvgIcon.CellPhone, '#929292'),
      onClick: () => this.sendCall(),
      showCompact: true,
      title: 'Mobile Phone'
    } as IContactPart,
    department: {
      icon: getSvg(SvgIcon.Department, '#929292'),
      showCompact: false,
      title: 'Department'
    } as IContactPart,
    title: {
      icon: getSvg(SvgIcon.Person, '#929292'),
      showCompact: false,
      title: 'Title'
    } as IContactPart,
    officeLocation: {
      icon: getSvg(SvgIcon.OfficeLocation, '#929292'),
      showCompact: true,
      title: 'Office Location'
    } as IContactPart
  };
  // tslint:enable: object-literal-sort-keys

  public constructor(person: User) {
    super();
    this._person = person;

    this._contactParts.email.value = getEmailFromGraphEntity(this._person);
    this._contactParts.chat.value = this._person.userPrincipalName;
    this._contactParts.cellPhone.value = this._person.mobilePhone;
    this._contactParts.department.value = this._person.department;
    this._contactParts.title.value = this._person.jobTitle;
    this._contactParts.officeLocation.value = this._person.officeLocation;

    if (this._person.businessPhones && this._person.businessPhones.length) {
      this._contactParts.businessPhone.value = this._person.businessPhones[0];
    }
  }

  /**
   * The name for display in the overview section.
   *
   * @readonly
   * @type {string}
   * @memberof MgtPersonCardContact
   */
  public get displayName(): string {
    return this.strings.contactSectionTitle;
  }

  // Defines the skeleton for what contact fields are available and what they do.

  /**
   * Render the icon for display in the navigation ribbon.
   *
   * @returns {TemplateResult}
   * @memberof MgtPersonCardContact
   */
  public renderIcon(): TemplateResult {
    return getSvg(SvgIcon.Contact);
  }

  /**
   * Reset any state in the section
   *
   * @protected
   * @memberof MgtPersonCardContact
   */
  public clearState() {
    super.clearState();
    for (const key of Object.keys(this._contactParts)) {
      this._contactParts[key].value = null;
    }
  }

  /**
   * Render the compact view
   *
   * @returns {TemplateResult}
   * @memberof MgtPersonCardContact
   */
  protected renderCompactView(): TemplateResult {
    let contentTemplate: TemplateResult;

    if (!this.hasData) {
      return null;
    }

    const availableParts: IContactPart[] = Object.values(this._contactParts).filter((p: IContactPart) => !!p.value);

    // Filter for compact mode parts with values
    let compactParts: IContactPart[] = Object.values(availableParts).filter(
      (p: IContactPart) => !!p.value && p.showCompact
    );

    if (!compactParts || !compactParts.length) {
      compactParts = Object.values(availableParts).slice(0, 2);
    }

    contentTemplate = html`
      ${compactParts.map(p => this.renderContactPart(p))}
    `;

    return html`
      <div class="root compact" dir=${this.direction}>
        ${contentTemplate}
      </div>
    `;
  }

  /**
   * Render the full view
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCardContact
   */
  protected renderFullView(): TemplateResult {
    let contentTemplate: TemplateResult;

    if (this.hasData) {
      // Filter for parts with values only
      const availableParts: IContactPart[] = Object.values(this._contactParts).filter((p: IContactPart) => !!p.value);
      contentTemplate = html`
        ${availableParts.map(part => this.renderContactPart(part))}
      `;
    }

    return html`
      <div class="root" dir=${this.direction}>
        <div class="title">${this.displayName}</div>
        ${contentTemplate}
      </div>
    `;
  }

  /**
   * Render a specific contact part
   *
   * @protected
   * @param {IContactPart} part
   * @returns {TemplateResult}
   * @memberof MgtPersonCardContact
   */
  protected renderContactPart(part: IContactPart): TemplateResult {
    let isPhone = false;

    if (part.title === 'Mobile Phone' || part.title === 'Business Phone') {
      isPhone = true;
    }

    const partLinkClasses = {
      part__link: true,
      phone: isPhone
    };

    const valueTemplate = part.onClick
      ? html`
          <span class=${classMap(partLinkClasses)} @click=${(e: Event) => part.onClick(e)}>${part.value}</span>
        `
      : html`
          ${part.value}
        `;

    return html`
      <div class="part" @click=${(e: MouseEvent) => this.handlePartClick(e, part.value)}>
        <div class="part__icon">${part.icon}</div>
        <div class="part__details">
          <div class="part__title">${part.title}</div>
          <div class="part__value">${valueTemplate}</div>
        </div>
        <div class="part__copy">
          ${getSvg(SvgIcon.Copy)}
        </div>
      </div>
    `;
  }

  /**
   * Handle the click event for contact parts
   *
   * @protected
   * @memberof MgtPersonCardContact
   */
  protected handlePartClick(e: MouseEvent, value: string): void {
    if (value) {
      navigator.clipboard.writeText(value);
    }
  }

  /**
   * Send a chat message to the user
   *
   * @protected
   * @memberof MgtPersonCardContact
   */
  protected sendChat(): void {
    const chat = this._contactParts.chat.value;
    if (!chat) {
      return;
    }

    const url = `https://teams.microsoft.com/l/chat/0/0?users=${chat}`;
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

  /**
   * Send an email to the user
   *
   * @protected
   * @memberof MgtPersonCardContact
   */
  protected sendEmail(): void {
    const email = this._contactParts.email.value;
    if (email) {
      window.open('mailto:' + email, '_blank');
    }
  }

  /**
   * Send a call to the user
   *
   * @protected
   * @memberof MgtPersonCardContact
   */
  protected sendCall(): void {
    const cellPhone = this._contactParts.cellPhone.value;
    if (cellPhone) {
      window.open('tel:' + cellPhone, '_blank');
    }
  }
}
