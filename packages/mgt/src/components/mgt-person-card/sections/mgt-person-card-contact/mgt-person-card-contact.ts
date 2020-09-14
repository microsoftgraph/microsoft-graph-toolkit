/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

// import * as GraphTypes from '@microsoft/microsoft-graph-types';
import * as GraphTypes from '@microsoft/microsoft-graph-types-beta';
import { customElement, html, TemplateResult } from 'lit-element';
import { getEmailFromGraphEntity } from '../../../../graph/graph.people';
import { TeamsHelper } from '../../../../utils/TeamsHelper';
import { BasePersonCardSection } from '../BasePersonCardSection';
import { styles } from './mgt-person-card-contact-css';
import { SvgIcon, getSvg } from '../../../../utils/SvgHelper';

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
 * The collection of relevant contact parts
 *
 * @interface IContactPartCollection
 */
interface IContactPartCollection {
  // tslint:disable-next-line: completed-docs
  cellPhone: IContactPart;
  // tslint:disable-next-line: completed-docs
  chat: IContactPart;
  // tslint:disable-next-line: completed-docs
  department: IContactPart;
  // tslint:disable-next-line: completed-docs
  email: IContactPart;
  // tslint:disable-next-line: completed-docs
  officeLocation: IContactPart;
  // tslint:disable-next-line: completed-docs
  title: IContactPart;
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

  /**
   * The name for display in the overview section.
   *
   * @readonly
   * @type {string}
   * @memberof MgtPersonCardContact
   */
  public get displayName(): string {
    return 'Contact';
  }

  // Defines the skeleton for what contact fields are available and what they do.

  // tslint:disable: object-literal-sort-keys
  private _contactParts: IContactPartCollection = {
    email: {
      icon: getSvg(SvgIcon.Email, '#929292'),
      onClick: () => this.sendEmail(),
      showCompact: true,
      title: 'Email'
    },
    chat: {
      icon: getSvg(SvgIcon.Chat, '#929292'),
      onClick: () => this.sendChat(),
      showCompact: false,
      title: 'Teams'
    },
    cellPhone: {
      icon: getSvg(SvgIcon.CellPhone, '#929292'),
      onClick: () => this.sendCall(),
      showCompact: true,
      title: 'Cell Phone'
    },
    department: {
      icon: getSvg(SvgIcon.Department, '#929292'),
      showCompact: false,
      title: 'Department'
    },
    title: {
      icon: getSvg(SvgIcon.Person, '#929292'),
      showCompact: false,
      title: 'Title'
    },
    officeLocation: {
      icon: getSvg(SvgIcon.OfficeLocation, '#929292'),
      showCompact: true,
      title: 'Office Location'
    }
  };
  // tslint:enable: object-literal-sort-keys

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

    if (this.isLoadingState) {
      contentTemplate = this.renderLoading();
    } else {
      // Filter for compact mode parts with values
      const compactParts: IContactPart[] = Object.values(this._contactParts).filter(
        (p: IContactPart) => !!p.value && p.showCompact
      );

      if (!compactParts || !compactParts.length) {
        contentTemplate = this.renderNoData();
      } else {
        contentTemplate = html`
          ${compactParts.map(p => this.renderContactPart(p))}
        `;
      }
    }

    return html`
      <div class="root compact">
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

    if (this.isLoadingState) {
      contentTemplate = this.renderLoading();
    } else if (!this._contactParts) {
      contentTemplate = this.renderNoData();
    } else {
      // Filter for parts with values only
      const availableParts: IContactPart[] = Object.values(this._contactParts).filter((p: IContactPart) => !!p.value);
      if (!availableParts.length) {
        contentTemplate = this.renderNoData();
      } else {
        contentTemplate = html`
          ${availableParts.map(part => this.renderContactPart(part))}
        `;
      }
    }

    return html`
      <div class="root">
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
    const valueTemplate = part.onClick
      ? html`
          <span class="part__link" @click=${(e: Event) => part.onClick(e)}>${part.value}</span>
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
    if (value && !this.isCompact) {
      navigator.clipboard.writeText(value);
    }
  }

  /**
   * Load the section state
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

    this._contactParts.email.value = getEmailFromGraphEntity(this.personDetails);
    this._contactParts.chat.value = personPerson.userPrincipalName;
    this._contactParts.cellPhone.value = userPerson.mobilePhone;
    this._contactParts.department.value = this.personDetails.department;
    this._contactParts.title.value = this.personDetails.jobTitle;
    this._contactParts.officeLocation.value = this.personDetails.officeLocation;
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
