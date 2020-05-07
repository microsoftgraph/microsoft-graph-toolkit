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
      <svg width="20" height="20" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path
          fill-rule="evenodd"
          clip-rule="evenodd"
          d="M11.3315 11.7855L11.3452 11.7913L11.3591 11.7967C11.4057 11.8149 11.4421 11.8389 11.4762 11.8729L12.8836 13.2797C12.9177 13.3137 12.9416 13.35 12.9597 13.3965L12.9651 13.4104L12.971 13.4241C12.9902 13.4693 13 13.5151 13 13.5699C13 13.6248 12.9901 13.6759 12.9676 13.7303L12.9675 13.7302L12.9628 13.7419C12.9458 13.7844 12.9218 13.822 12.8836 13.8602L12.7941 13.9497C12.57 14.1736 12.3673 14.37 12.1855 14.5396C12.0445 14.6712 11.9037 14.7767 11.7639 14.8601C11.6429 14.9284 11.499 14.9865 11.3265 15.0296C11.1716 15.0683 10.9523 15.0957 10.6541 15.0957C10.226 15.0957 9.77339 15.0291 9.29316 14.8873C8.79837 14.7411 8.29897 14.537 7.79466 14.2716C7.29099 14.0043 6.7904 13.6843 6.2931 13.3097C5.79797 12.9334 5.32815 12.5234 4.88351 12.0793C4.44215 11.6299 4.03516 11.1575 3.66219 10.6619C3.2916 10.1652 2.97561 9.66531 2.71228 9.16232C2.44987 8.65684 2.24939 8.16115 2.1072 7.67496C1.96889 7.20203 1.9043 6.75953 1.9043 6.34426C1.9043 6.04676 1.92945 5.8283 1.96505 5.67477C2.00758 5.51024 2.06491 5.36965 2.13346 5.24839C2.21771 5.10725 2.32161 4.96874 2.44802 4.83312C2.61763 4.65148 2.814 4.44899 3.03788 4.22522L3.14024 4.12291C3.18256 4.08061 3.22694 4.05101 3.2785 4.0291C3.32105 4.01103 3.36887 4 3.43131 4C3.48634 4 3.53235 4.00983 3.57773 4.0291L3.57771 4.02914L3.58585 4.03251C3.6411 4.05541 3.68414 4.08469 3.72239 4.12291L5.12984 5.52968C5.16387 5.56369 5.18781 5.60004 5.20594 5.64649L5.21136 5.66038L5.2172 5.6741C5.23642 5.71931 5.24622 5.76513 5.24622 5.81992C5.24622 5.88591 5.23631 5.9178 5.23225 5.92835C5.20132 5.99698 5.16601 6.05438 5.12738 6.10352C5.07447 6.17083 5.01613 6.23146 4.95179 6.28625L4.94467 6.2923L4.93768 6.29849C4.81323 6.40852 4.69584 6.51839 4.58605 6.62811C4.45174 6.76237 4.32909 6.90895 4.22378 7.06908C4.0535 7.32254 3.96032 7.61696 3.96032 7.93007C3.96032 8.35888 4.12439 8.75141 4.42612 9.05299L7.95115 12.5763C8.25285 12.8779 8.64535 13.0417 9.07392 13.0417C9.38687 13.0417 9.68121 12.9486 9.9347 12.7785C10.0949 12.6733 10.2415 12.5507 10.3758 12.4164C10.4856 12.3067 10.5955 12.1894 10.7056 12.065L10.7118 12.058L10.7179 12.0509C10.7727 11.9866 10.8333 11.9283 10.9007 11.8754C10.9499 11.8367 11.0074 11.8014 11.0761 11.7704C11.0868 11.7663 11.1189 11.7564 11.1851 11.7564C11.2401 11.7564 11.2861 11.7662 11.3315 11.7855ZM11.5689 15.9998C11.8248 15.9358 12.0573 15.8442 12.2663 15.7248C12.4753 15.6012 12.6757 15.4499 12.8676 15.2708C13.0596 15.0918 13.2707 14.8872 13.501 14.657L13.5906 14.5674C13.7228 14.4353 13.823 14.284 13.8912 14.1134C13.9637 13.9387 14 13.7575 14 13.5699C14 13.3824 13.9637 13.2033 13.8912 13.0328C13.823 12.858 13.7228 12.7045 13.5906 12.5724L12.1831 11.1656C12.0509 11.0335 11.8974 10.9333 11.7225 10.8651C11.5519 10.7926 11.3728 10.7564 11.1851 10.7564C10.9974 10.7564 10.829 10.7884 10.6797 10.8523C10.5347 10.9163 10.4025 10.9951 10.283 11.0889C10.1636 11.1827 10.0549 11.2871 9.95677 11.4022C9.85868 11.5131 9.76272 11.6154 9.66889 11.7092C9.57506 11.8029 9.47909 11.8818 9.381 11.9457C9.28717 12.0097 9.18481 12.0417 9.07392 12.0417C8.91185 12.0417 8.77323 11.9841 8.65808 11.869L5.13305 8.34571C5.0179 8.23061 4.96032 8.09206 4.96032 7.93007C4.96032 7.81924 4.99231 7.71693 5.05628 7.62314C5.12026 7.52509 5.19916 7.42918 5.29299 7.33539C5.38682 7.24161 5.48918 7.14569 5.60007 7.04765C5.71522 6.9496 5.81972 6.84089 5.91355 6.72153C6.00738 6.60217 6.08628 6.47002 6.15026 6.32508C6.21423 6.17588 6.24622 6.00749 6.24622 5.81992C6.24622 5.63236 6.20997 5.45331 6.13746 5.2828C6.06922 5.10802 5.96899 4.95455 5.83678 4.8224L4.42932 3.41564C4.29711 3.28348 4.14357 3.18117 3.9687 3.1087C3.7981 3.03623 3.61897 3 3.43131 3C3.23939 3 3.05813 3.03623 2.88752 3.1087C2.71692 3.18117 2.56552 3.28348 2.4333 3.41564L2.33094 3.51795C2.10063 3.74814 1.89591 3.95916 1.71678 4.15099C1.54192 4.33856 1.39264 4.53678 1.26895 4.74567C1.14953 4.95455 1.05784 5.18475 0.993862 5.43626C0.934152 5.68777 0.904297 5.99044 0.904297 6.34426C0.904297 6.86434 0.985332 7.40147 1.1474 7.95565C1.30947 8.50983 1.53552 9.06614 1.82554 9.62458C2.11556 10.1788 2.46102 10.7244 2.86193 11.2615C3.26285 11.7944 3.70001 12.3017 4.17342 12.7834C4.65111 13.2609 5.15651 13.7021 5.68963 14.107C6.22703 14.512 6.77295 14.8616 7.3274 15.1557C7.88611 15.4499 8.44696 15.6801 9.00994 15.8463C9.57292 16.0126 10.121 16.0957 10.6541 16.0957C11.0081 16.0957 11.313 16.0637 11.5689 15.9998ZM14.9616 14H17.0002C17.5525 14 18.0002 13.5523 18.0002 13V5C18.0002 4.44772 17.5525 4 17.0002 4H6.43555L6.54385 4.11514C6.77073 4.34192 6.94536 4.60775 7.06365 4.90525C7.07677 4.93665 7.08919 4.96824 7.10089 5H16.8821L10.2238 8.32913C10.083 8.39951 9.91736 8.39951 9.7766 8.32913L7.22381 7.05273L6.7766 7.94716L9.32938 9.22355C9.75167 9.4347 10.2487 9.4347 10.671 9.22356L17.0002 6.05895V13H14.9314C14.9772 13.1858 15.0001 13.3764 15.0001 13.5699C15.0001 13.7154 14.9872 13.8589 14.9616 14Z"
          fill="#605E5C"
        />
      </svg>
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
