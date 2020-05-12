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
import { TeamsHelper } from '../../../../../utils/TeamsHelper';
import { BasePersonCardSection } from '../BasePersonCardSection';
import { styles } from './mgt-person-card-contact-css';

/**
 * foo
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
 * foo
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

  // Defines the skeleton for what contact fields are available and what they do.

  // tslint:disable: object-literal-sort-keys
  private _contactParts: IContactPartCollection = {
    email: {
      icon: icons.email,
      onClick: () => this.sendEmail(),
      showCompact: true,
      title: 'Email'
    },
    chat: {
      icon: icons.chat,
      onClick: () => this.sendChat(),
      showCompact: false,
      title: 'Teams'
    },
    cellPhone: {
      icon: icons.cellPhone,
      onClick: () => this.sendCall(),
      showCompact: true,
      title: 'Cell Phone'
    },
    department: {
      icon: icons.department,
      showCompact: false,
      title: 'Department'
    },
    title: {
      icon: icons.title,
      showCompact: false,
      title: 'Title'
    },
    officeLocation: {
      icon: icons.officeLocation,
      onClick: () => this.showOfficeLocation(),
      showCompact: true,
      title: 'Office Location'
    }
  };
  // tslint:enable: object-literal-sort-keys

  /**
   * foo
   *
   * @returns {TemplateResult}
   * @memberof MgtPersonCardContact
   */
  public renderIcon(): TemplateResult {
    return html`
      <svg xmlns="http://www.w3.org/2000/svg">
        <path
          d="M11.3315 11.7855L11.3452 11.7913L11.3591 11.7967C11.4057 11.8149 11.4421 11.8389 11.4762 11.8729L12.8836 13.2797C12.9177 13.3137 12.9416 13.35 12.9597 13.3965L12.9651 13.4104L12.971 13.4241C12.9902 13.4693 13 13.5151 13 13.5699C13 13.6248 12.9901 13.6759 12.9676 13.7303L12.9675 13.7302L12.9628 13.7419C12.9458 13.7844 12.9218 13.822 12.8836 13.8602L12.7941 13.9497C12.57 14.1736 12.3673 14.37 12.1855 14.5396C12.0445 14.6712 11.9037 14.7767 11.7639 14.8601C11.6429 14.9284 11.499 14.9865 11.3265 15.0296C11.1716 15.0683 10.9523 15.0957 10.6541 15.0957C10.226 15.0957 9.77339 15.0291 9.29316 14.8873C8.79837 14.7411 8.29897 14.537 7.79466 14.2716C7.29099 14.0043 6.7904 13.6843 6.2931 13.3097C5.79797 12.9334 5.32815 12.5234 4.88351 12.0793C4.44215 11.6299 4.03516 11.1575 3.66219 10.6619C3.2916 10.1652 2.97561 9.66531 2.71228 9.16232C2.44987 8.65684 2.24939 8.16115 2.1072 7.67496C1.96889 7.20203 1.9043 6.75953 1.9043 6.34426C1.9043 6.04676 1.92945 5.8283 1.96505 5.67477C2.00758 5.51024 2.06491 5.36965 2.13346 5.24839C2.21771 5.10725 2.32161 4.96874 2.44802 4.83312C2.61763 4.65148 2.814 4.44899 3.03788 4.22522L3.14024 4.12291C3.18256 4.08061 3.22694 4.05101 3.2785 4.0291C3.32105 4.01103 3.36887 4 3.43131 4C3.48634 4 3.53235 4.00983 3.57773 4.0291L3.57771 4.02914L3.58585 4.03251C3.6411 4.05541 3.68414 4.08469 3.72239 4.12291L5.12984 5.52968C5.16387 5.56369 5.18781 5.60004 5.20594 5.64649L5.21136 5.66038L5.2172 5.6741C5.23642 5.71931 5.24622 5.76513 5.24622 5.81992C5.24622 5.88591 5.23631 5.9178 5.23225 5.92835C5.20132 5.99698 5.16601 6.05438 5.12738 6.10352C5.07447 6.17083 5.01613 6.23146 4.95179 6.28625L4.94467 6.2923L4.93768 6.29849C4.81323 6.40852 4.69584 6.51839 4.58605 6.62811C4.45174 6.76237 4.32909 6.90895 4.22378 7.06908C4.0535 7.32254 3.96032 7.61696 3.96032 7.93007C3.96032 8.35888 4.12439 8.75141 4.42612 9.05299L7.95115 12.5763C8.25285 12.8779 8.64535 13.0417 9.07392 13.0417C9.38687 13.0417 9.68121 12.9486 9.9347 12.7785C10.0949 12.6733 10.2415 12.5507 10.3758 12.4164C10.4856 12.3067 10.5955 12.1894 10.7056 12.065L10.7118 12.058L10.7179 12.0509C10.7727 11.9866 10.8333 11.9283 10.9007 11.8754C10.9499 11.8367 11.0074 11.8014 11.0761 11.7704C11.0868 11.7663 11.1189 11.7564 11.1851 11.7564C11.2401 11.7564 11.2861 11.7662 11.3315 11.7855ZM11.5689 15.9998C11.8248 15.9358 12.0573 15.8442 12.2663 15.7248C12.4753 15.6012 12.6757 15.4499 12.8676 15.2708C13.0596 15.0918 13.2707 14.8872 13.501 14.657L13.5906 14.5674C13.7228 14.4353 13.823 14.284 13.8912 14.1134C13.9637 13.9387 14 13.7575 14 13.5699C14 13.3824 13.9637 13.2033 13.8912 13.0328C13.823 12.858 13.7228 12.7045 13.5906 12.5724L12.1831 11.1656C12.0509 11.0335 11.8974 10.9333 11.7225 10.8651C11.5519 10.7926 11.3728 10.7564 11.1851 10.7564C10.9974 10.7564 10.829 10.7884 10.6797 10.8523C10.5347 10.9163 10.4025 10.9951 10.283 11.0889C10.1636 11.1827 10.0549 11.2871 9.95677 11.4022C9.85868 11.5131 9.76272 11.6154 9.66889 11.7092C9.57506 11.8029 9.47909 11.8818 9.381 11.9457C9.28717 12.0097 9.18481 12.0417 9.07392 12.0417C8.91185 12.0417 8.77323 11.9841 8.65808 11.869L5.13305 8.34571C5.0179 8.23061 4.96032 8.09206 4.96032 7.93007C4.96032 7.81924 4.99231 7.71693 5.05628 7.62314C5.12026 7.52509 5.19916 7.42918 5.29299 7.33539C5.38682 7.24161 5.48918 7.14569 5.60007 7.04765C5.71522 6.9496 5.81972 6.84089 5.91355 6.72153C6.00738 6.60217 6.08628 6.47002 6.15026 6.32508C6.21423 6.17588 6.24622 6.00749 6.24622 5.81992C6.24622 5.63236 6.20997 5.45331 6.13746 5.2828C6.06922 5.10802 5.96899 4.95455 5.83678 4.8224L4.42932 3.41564C4.29711 3.28348 4.14357 3.18117 3.9687 3.1087C3.7981 3.03623 3.61897 3 3.43131 3C3.23939 3 3.05813 3.03623 2.88752 3.1087C2.71692 3.18117 2.56552 3.28348 2.4333 3.41564L2.33094 3.51795C2.10063 3.74814 1.89591 3.95916 1.71678 4.15099C1.54192 4.33856 1.39264 4.53678 1.26895 4.74567C1.14953 4.95455 1.05784 5.18475 0.993862 5.43626C0.934152 5.68777 0.904297 5.99044 0.904297 6.34426C0.904297 6.86434 0.985332 7.40147 1.1474 7.95565C1.30947 8.50983 1.53552 9.06614 1.82554 9.62458C2.11556 10.1788 2.46102 10.7244 2.86193 11.2615C3.26285 11.7944 3.70001 12.3017 4.17342 12.7834C4.65111 13.2609 5.15651 13.7021 5.68963 14.107C6.22703 14.512 6.77295 14.8616 7.3274 15.1557C7.88611 15.4499 8.44696 15.6801 9.00994 15.8463C9.57292 16.0126 10.121 16.0957 10.6541 16.0957C11.0081 16.0957 11.313 16.0637 11.5689 15.9998ZM14.9616 14H17.0002C17.5525 14 18.0002 13.5523 18.0002 13V5C18.0002 4.44772 17.5525 4 17.0002 4H6.43555L6.54385 4.11514C6.77073 4.34192 6.94536 4.60775 7.06365 4.90525C7.07677 4.93665 7.08919 4.96824 7.10089 5H16.8821L10.2238 8.32913C10.083 8.39951 9.91736 8.39951 9.7766 8.32913L7.22381 7.05273L6.7766 7.94716L9.32938 9.22355C9.75167 9.4347 10.2487 9.4347 10.671 9.22356L17.0002 6.05895V13H14.9314C14.9772 13.1858 15.0001 13.3764 15.0001 13.5699C15.0001 13.7154 14.9872 13.8589 14.9616 14Z"
        />
      </svg>
    `;
  }

  /**
   * foo
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
   * foo
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
   * foo
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
   * foo
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCardContact
   */
  protected renderLoading(): TemplateResult {
    return html`
      <div class="loading">Loading</div>
    `;
  }

  /**
   * foo
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCardContact
   */
  protected renderNoData(): TemplateResult {
    return html`
      <div class="no-data">No data</div>
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
          <svg width="13" height="14" viewBox="0 0 13 14" xmlns="http://www.w3.org/2000/svg">
            <path
              d="M12.625 5.50293V14H3.875V11.375H0.375V0H6.24707L8.87207 2.625H9.74707L12.625 5.50293ZM10 5.25H11.1279L10 4.12207V5.25ZM3.875 2.625H7.62793L5.87793 0.875H1.25V10.5H3.875V2.625ZM11.75 6.125H9.125V3.5H4.75V13.125H11.75V6.125Z"
            />
          </svg>
        </div>
      </div>
    `;
  }

  /**
   * foo
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

    this._contactParts.email.value = getEmailFromGraphEntity(this.personDetails);
    this._contactParts.chat.value = personPerson.userPrincipalName;
    this._contactParts.cellPhone.value = userPerson.mobilePhone;
    this._contactParts.department.value = this.personDetails.department;
    this._contactParts.title.value = this.personDetails.jobTitle;
    this._contactParts.officeLocation.value = this.personDetails.officeLocation;
  }

  /**
   * foo
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
   * foo
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
   * foo
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

  /**
   * foo
   *
   * @protected
   * @memberof MgtPersonCardContact
   */
  protected showOfficeLocation(): void {
    const officeLocation = this._contactParts.officeLocation.value;
    if (!officeLocation) {
      return;
    }
    // TODO: Show the office location somehow.
  }
}

/**
 * Icon templates
 */
const icons = {
  cellPhone: html`
    <svg width="10" height="15" viewBox="0 0 10 15" fill="none" xmlns="http://www.w3.org/2000/svg">
      <rect x="0.5" y="0.5" width="9" height="14" rx="0.9" stroke="#929292" />
      <rect x="3" y="12" width="4" height="1" rx="0.5" fill="#929292" />
    </svg>
  `,
  chat: html`
    <svg width="17" height="15" viewBox="0 0 17 15" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path
        d="M2.67824 0C2.3824 0 2.14258 0.239826 2.14258 0.535665C2.14258 0.831505 2.3824 1.07133 2.67824 1.07133H13.9985C14.5508 1.07133 14.9985 1.51905 14.9985 2.07133V12.3203C14.9985 12.6161 15.2384 12.856 15.5342 12.856C15.8301 12.856 16.0699 12.6161 16.0699 12.3203V2C16.0699 0.895431 15.1744 0 14.0699 0H2.67824Z"
        fill="#929292"
      />
      <path
        d="M9.34097 11.4769L9.3309 11.4769L9.32085 11.4773C6.74142 11.5804 3.51639 11.5801 1.6855 11.1657L1.6855 11.1657L1.67972 11.1644C1.30373 11.084 0.937799 10.8292 0.816663 10.7209L0.81023 10.7152L0.803602 10.7096C0.601843 10.5414 0.5 10.3403 0.5 10.0423V3.55978C0.5 3.07185 0.912353 2.64258 1.47765 2.64258H12.4497C13.0149 2.64258 13.4273 3.07185 13.4273 3.55978V13.9745C13.4273 14.1578 13.2827 14.3397 13.067 14.4366C12.8614 14.5288 12.6204 14.5275 12.457 14.3904L10.5932 11.9699C10.5218 11.8587 10.4457 11.7646 10.3579 11.6904C10.2528 11.6015 10.1556 11.5622 10.0869 11.5401L9.93419 12.0163L10.0869 11.5401C9.89097 11.4773 9.65413 11.477 9.34097 11.4769Z"
        stroke="#929292"
      />
    </svg>
  `,
  department: html`
    <svg width="14" height="14" viewBox="0 0 14 14" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path
        d="M9.625 3.5H14V11.375H0V3.5H4.375V1.75H9.625V3.5ZM5.25 2.625V3.5H8.75V2.625H5.25ZM13.125 4.375H0.875V6.125H3.5V5.25H4.375V6.125H9.625V5.25H10.5V6.125H13.125V4.375ZM0.875 10.5H13.125V7H10.5V7.875H9.625V7H4.375V7.875H3.5V7H0.875V10.5Z"
        fill="#929292"
      />
    </svg>
  `,
  email: html`
    <svg width="14" height="10" viewBox="0 0 14 10" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path
        d="M0.5 0.772727C0.5 0.622104 0.622104 0.5 0.772727 0.5H13.2273C13.3779 0.5 13.5 0.622104 13.5 0.772727V9.22727C13.5 9.3779 13.3779 9.5 13.2273 9.5H0.772727C0.622104 9.5 0.5 9.3779 0.5 9.22727V0.772727Z"
        stroke="#929292"
      />
      <path d="M13.5 0.5L7.18923 4.70314C6.92113 4.8817 6.57039 4.87522 6.30907 4.68687L0.5 0.5" stroke="#929292" />
    </svg>
  `,
  officeLocation: html`
    <svg width="14" height="17" viewBox="0 0 14 17" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path
        d="M6.78489 16.3832L6.78263 16.3859C6.75278 16.4216 6.71543 16.4503 6.67324 16.4701L6.88498 16.923L6.67324 16.4701C6.63105 16.4898 6.58504 16.5 6.53846 16.5C6.49188 16.5 6.44588 16.4898 6.40368 16.4701C6.36149 16.4503 6.32415 16.4216 6.2943 16.3859L6.29202 16.3832C5.47882 15.4241 4.01597 13.6289 2.75914 11.7172C2.13055 10.7611 1.56021 9.78597 1.14862 8.87887C0.732553 7.96189 0.5 7.15987 0.5 6.53687C0.5 3.20251 3.20343 0.5 6.53846 0.5C9.87349 0.5 12.5769 3.20251 12.5769 6.53687C12.5769 7.16011 12.3444 7.96225 11.9283 8.87925C11.5167 9.78639 10.9464 10.7615 10.3178 11.7175C9.06097 13.6291 7.59812 15.424 6.78489 16.3832Z"
        stroke="#929292"
      />
      <path
        d="M4.40039 6.53921C4.40039 5.37092 5.34748 4.42383 6.51577 4.42383C7.68407 4.42383 8.63116 5.37092 8.63116 6.53921C8.63116 7.70751 7.68407 8.6546 6.51577 8.6546C5.34748 8.6546 4.40039 7.70751 4.40039 6.53921Z"
        stroke="#929292"
      />
    </svg>
  `,
  title: html`
    <svg width="12" height="12" viewBox="0 0 12 12" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path
        fill-rule="evenodd"
        clip-rule="evenodd"
        d="M8.91699 4C8.91699 5.65685 7.57385 7 5.91699 7C4.26014 7 2.91699 5.65685 2.91699 4C2.91699 2.34315 4.26014 1 5.91699 1C7.57385 1 8.91699 2.34315 8.91699 4ZM8.04431 7.38803C9.16935 6.68014 9.91699 5.42738 9.91699 4C9.91699 1.79086 8.12613 0 5.91699 0C3.70785 0 1.91699 1.79086 1.91699 4C1.91699 5.42739 2.66465 6.68016 3.78972 7.38805C1.82681 8.13254 0.356122 9.8773 0 12H1.01706C1.48033 9.71776 3.49808 8 5.91704 8C8.336 8 10.3538 9.71776 10.817 12H11.8341C11.478 9.87728 10.0072 8.1325 8.04431 7.38803Z"
        fill="#929292"
      />
    </svg>
  `
};
