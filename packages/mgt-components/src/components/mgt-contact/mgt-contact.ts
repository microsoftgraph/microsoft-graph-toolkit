/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { User } from '@microsoft/microsoft-graph-types';
import { html, TemplateResult } from 'lit';
import { TeamsHelper, customElement } from '@microsoft/mgt-element';
import { classMap } from 'lit/directives/class-map.js';

import { getEmailFromGraphEntity } from '../../graph/graph.people';
import { BasePersonCardSection } from '../BasePersonCardSection';
import { styles } from './mgt-contact-css';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';
import { strings } from './strings';

/**
 * Represents a contact part and its metadata
 *
 * @interface IContactPart
 */
interface IContactPart {
  icon: TemplateResult;
  title: string;
  value?: string;
  onClick?: (e: Event) => void;
  showCompact: boolean;
}

type Protocol = 'mailto:' | 'tel:';

/**
 * The contact details subsection of the person card
 *
 * @export
 * @class MgtContact
 * @extends {MgtTemplatedComponent}
 */
@customElement('contact')
// @customElement('mgt-contact')
export class MgtContact extends BasePersonCardSection {
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

  private readonly _person?: User;

  private readonly _contactParts: Record<string, IContactPart> = {
    email: {
      icon: getSvg(SvgIcon.Email),
      onClick: () => this.sendEmail(getEmailFromGraphEntity(this._person)),
      showCompact: true,
      title: this.strings.emailTitle
    },
    chat: {
      icon: getSvg(SvgIcon.Chat),
      onClick: () => this.sendChat(this._person?.userPrincipalName),
      showCompact: false,
      title: this.strings.chatTitle
    },
    businessPhone: {
      icon: getSvg(SvgIcon.Phone),
      onClick: () => this.sendCall(this._person?.businessPhones?.length > 0 ? this._person.businessPhones[0] : null),
      showCompact: true,
      title: this.strings.businessPhoneTitle
    },
    cellPhone: {
      icon: getSvg(SvgIcon.CellPhone),
      onClick: () => this.sendCall(this._person?.mobilePhone),
      showCompact: true,
      title: this.strings.cellPhoneTitle
    },
    department: {
      icon: getSvg(SvgIcon.Department),
      showCompact: false,
      title: this.strings.departmentTitle
    },
    title: {
      icon: getSvg(SvgIcon.Person),
      showCompact: false,
      title: this.strings.titleTitle
    },
    officeLocation: {
      icon: getSvg(SvgIcon.OfficeLocation),
      showCompact: true,
      title: this.strings.officeLocationTitle
    }
  };

  public constructor(person: User) {
    super();
    this._person = person;

    this._contactParts.email.value = getEmailFromGraphEntity(this._person);
    this._contactParts.chat.value = this._person.userPrincipalName;
    this._contactParts.cellPhone.value = this._person.mobilePhone;
    this._contactParts.department.value = this._person.department;
    this._contactParts.title.value = this._person.jobTitle;
    this._contactParts.officeLocation.value = this._person.officeLocation;

    if (this._person.businessPhones?.length) {
      this._contactParts.businessPhone.value = this._person.businessPhones[0];
    }
  }

  /**
   * The name for display in the overview section.
   *
   * @readonly
   * @type {string}
   * @memberof MgtContact
   */
  public get displayName(): string {
    return this.strings.contactSectionTitle;
  }

  /**
   * The title for display when rendered as a full card.
   *
   * @readonly
   * @type {string}
   * @memberof MgtContact
   */
  public get cardTitle(): string {
    return this.strings.contactSectionTitle;
  }

  // Defines the skeleton for what contact fields are available and what they do.

  /**
   * Render the icon for display in the navigation ribbon.
   *
   * @returns {TemplateResult}
   * @memberof MgtContact
   */
  public renderIcon(): TemplateResult {
    return getSvg(SvgIcon.Contact);
  }

  /**
   * Reset any state in the section
   *
   * @protected
   * @memberof MgtContact
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
   * @memberof MgtContact
   */
  protected renderCompactView(): TemplateResult {
    if (!this.hasData) {
      return null;
    }

    const availableParts: IContactPart[] = Object.values(this._contactParts).filter((p: IContactPart) => !!p.value);

    // Filter for compact mode parts with values
    let compactParts: IContactPart[] = Object.values(availableParts).filter(
      (p: IContactPart) => !!p.value && p.showCompact
    );

    if (!compactParts?.length) {
      compactParts = Object.values(availableParts).slice(0, 2);
    }

    const contentTemplate = html`
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
   * @memberof MgtContact
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
   * @memberof MgtContact
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
      <div class="part" role="button" @click=${(e: MouseEvent) => this.handlePartClick(e, part.value)} tabindex="0">
        <div class="part__icon" aria-label=${part.title} title=${part.title}>${part.icon}</div>
        <div class="part__details">
          <div class="part__title">${part.title}</div>
          <div class="part__value" title=${part.title}>${valueTemplate}</div>
        </div>
        <div
          class="part__copy"
          aria-label=${this.strings.copyToClipboardButton}
          title=${this.strings.copyToClipboardButton}
        >
          ${getSvg(SvgIcon.Copy)}
        </div>
      </div>
    `;
  }

  /**
   * Handle the click event for contact parts
   *
   * @protected
   * @memberof MgtContact
   */
  protected handlePartClick(e: MouseEvent, value: string): void {
    if (value) {
      void navigator.clipboard.writeText(value);
    }
  }

  private sendLink(protocol: Protocol, resource: string): void {
    if (resource) {
      window.open(`${protocol}${resource}`, '_blank', 'noreferrer');
    } else {
      // eslint-disable-next-line no-console
      console.error(`🦒: Target resource for ${protocol} link was not provided: resource: ${resource}`);
    }
  }

  /**
   * Send a chat message to the user
   *
   * @protected
   * @memberof MgtContact
   */
  protected sendChat(upn: string): void {
    if (!upn) {
      // eslint-disable-next-line no-console
      console.error("🦒: Can't send chat when upn is not provided");
      return;
    }

    const url = `https://teams.microsoft.com/l/chat/0/0?users=${upn}`;
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

  /**
   * Send an email to the user
   *
   * @protected
   * @memberof MgtContact
   */
  protected sendEmail(email: string): void {
    this.sendLink('mailto:', email);
  }

  /**
   * Send a call to the user
   *
   * @protected
   * @memberof MgtContact
   */
  protected sendCall = (phone: string): void => {
    this.sendLink('tel:', phone);
  };
}
