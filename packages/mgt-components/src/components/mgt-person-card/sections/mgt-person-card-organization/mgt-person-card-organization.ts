/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { User } from '@microsoft/microsoft-graph-types';
import { customElement, html, TemplateResult } from 'lit-element';

import { BasePersonCardSection } from '../BasePersonCardSection';
import { getSvg, SvgIcon } from '../../../../utils/SvgHelper';
import { MgtPersonCardState } from '../../mgt-person-card.types';
import { styles } from './mgt-person-card-organization-css';
import { strings } from './strings';
import { ViewType } from '../../../../graph/types';

/**
 * The member organization subsection of the person card
 *
 * @export
 * @class MgtPersonCardProfile
 * @extends {MgtTemplatedComponent}
 */
@customElement('mgt-person-card-organization')
export class MgtPersonCardOrganization extends BasePersonCardSection {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * returns component strings
   *
   * @readonly
   * @protected
   * @memberof MgtBaseComponent
   */
  protected get strings() {
    return strings;
  }

  private _state: MgtPersonCardState;
  private _me: User;

  constructor(state: MgtPersonCardState, me: User) {
    super();
    this._state = state;
    this._me = me;
  }

  /**
   * Reset any state in the section
   *
   * @protected
   * @memberof MgtPersonCardMessages
   */
  public clearState(): void {
    super.clearState();
    this._state = null;
    this._me = null;
  }

  /**
   * The name for display in the overview section.
   *
   * @readonly
   * @type {string}
   * @memberof MgtPersonCardOrganization
   */
  public get displayName(): string {
    const { person, directReports } = this._state;

    if (!person.manager && directReports && directReports.length) {
      return `${this.strings.directReportsSectionTitle} (${directReports.length})`;
    }

    return this.strings.reportsToSectionTitle;
  }

  /**
   * Render the icon for display in the navigation ribbon.
   *
   * @returns {TemplateResult}
   * @memberof MgtPersonCardOrganization
   */
  public renderIcon(): TemplateResult {
    return getSvg(SvgIcon.Organization);
  }

  /**
   * Render the compact view
   *
   * @returns {TemplateResult}
   * @memberof MgtPersonCardOrganization
   */
  protected renderCompactView(): TemplateResult {
    let contentTemplate: TemplateResult;

    if (!this._state || !this._state.person) {
      return null;
    }

    const { person, directReports } = this._state;

    if (!person) {
      return null;
    } else if (person.manager) {
      contentTemplate = this.renderCoworker(person.manager);
    } else if (directReports && directReports.length) {
      contentTemplate = this.renderCompactDirectReports();
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
   * @memberof MgtPersonCardOrganization
   */
  protected renderFullView(): TemplateResult {
    let contentTemplate: TemplateResult;

    if (!this._state || !this._state.person) {
      return null;
    }

    const { person, directReports, people } = this._state;

    if (!person && !directReports && !people) {
      return null;
    } else {
      const managerTemplates = this.renderManagers();
      const currentUserTemplate = this.renderCurrentUser();
      const directReportsTemplate = this.renderDirectReports();
      const coworkersTemplate = this.renderCoworkers();

      contentTemplate = html`
        ${managerTemplates} ${currentUserTemplate} ${directReportsTemplate} ${coworkersTemplate}
      `;
    }

    return html`
      <div class="root" dir=${this.direction}>
        <div class="title" tabindex="0">${this.strings.organizationSectionTitle}</div>
        ${contentTemplate}
      </div>
    `;
  }

  /**
   * Render a manager org member
   *
   * @protected
   * @param {User} person
   * @returns {TemplateResult}
   * @memberof MgtPersonCardOrganization
   */
  protected renderManager(person: User): TemplateResult {
    return html`
      <div class="org-member" @keydown=${(e: KeyboardEvent) => {
        e.code === 'Enter' ? this.navigateCard(person) : '';
      }} @click=${() => this.navigateCard(person)}>
        <div class="org-member__person">
          <mgt-person
            .personDetails=${person}
            .line2Property=${'jobTitle'}
            .line3Property=${'department'}
            .fetchImage=${true}
            .view=${ViewType.threelines}
          ></mgt-person>
        </div>
        <div tabindex="0" class="org-member__more">
          ${getSvg(SvgIcon.ExpandRight)}
        </div>
      </div>
      <div class="org-member__separator"></div>
    `;
  }

  /**
   * Render a manager org member
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCardOrganization
   */
  protected renderManagers(): TemplateResult[] {
    const { person } = this._state;
    if (!person || !person.manager) {
      return null;
    }

    const managers = [];
    let currentManager = person;
    while (currentManager.manager) {
      managers.push(currentManager.manager);
      currentManager = currentManager.manager;
    }

    if (!managers.length) {
      return null;
    }

    return managers.reverse().map(manager => this.renderManager(manager));
  }

  /**
   * Render a direct report
   *
   * @protected
   * @param {User} person
   * @returns {TemplateResult}
   * @memberof MgtPersonCardOrganization
   */
  protected renderDirectReports(): TemplateResult {
    const { directReports } = this._state;
    if (!directReports || !directReports.length) {
      return null;
    }

    return html`
      <div class="org-member__separator"></div>
      <div>
        ${directReports.map(
          person => html`
            <div class="org-member org-member--direct-report" @keydown=${(e: KeyboardEvent) => {
              e.code === 'Enter' ? this.navigateCard(person) : '';
            }} @click=${() => this.navigateCard(person)}>
              <div class="org-member__person">
                <mgt-person
                  .personDetails=${person}
                  .line2Property=${'jobTitle'}
                  .line3Property=${'department'}
                  .fetchImage=${true}
                  .view=${ViewType.twolines}
                ></mgt-person>
              </div>
              <div tabindex="0" class="org-member__more">
                ${getSvg(SvgIcon.ExpandRight)}
              </div>
            </div>
          `
        )}
      </div>
    `;
  }

  /**
   * Render direct reports in compact view
   *
   * @protected
   * @param {User} person
   * @returns {TemplateResult}
   * @memberof MgtPersonCardOrganization
   */
  protected renderCompactDirectReports(): TemplateResult {
    const { directReports } = this._state;

    return html`
      <div class="direct-report__compact">
        ${directReports.slice(0, 6).map(
          person => html`
            <div class="direct-report" @keydown=${(e: KeyboardEvent) => {
              e.code === 'Enter' ? this.navigateCard(person) : '';
            }} @click=${() => this.navigateCard(person)} @keydown=${(e: KeyboardEvent) => {
            e.code === 'Enter' ? this.navigateCard(person) : '';
          }}>
              <mgt-person .personDetails=${person} .fetchImage=${true}></mgt-person>
            </div>
          `
        )}
      </div>
    `;
  }

  /**
   * Render the user/self member
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCardOrganization
   */
  protected renderCurrentUser(): TemplateResult {
    const { person } = this._state;
    return html`
      <div class="org-member org-member--target">
        <div class="org-member__person">
          <mgt-person
            .personDetails=${person}
            .line2Property=${'jobTitle'}
            .line3Property=${'department'}
            .fetchImage=${true}
            .view=${ViewType.threelines}
          ></mgt-person>
        </div>
      </div>
    `;
  }

  /**
   * Render a coworker org member
   *
   * @protected
   * @param {User} person
   * @returns {TemplateResult}
   * @memberof MgtPersonCardOrganization
   */
  protected renderCoworker(person: User): TemplateResult {
    return html`
      <div class="coworker" @keydown=${(e: KeyboardEvent) => {
        e.code === 'Enter' ? this.navigateCard(person) : '';
      }} @click=${() => this.navigateCard(person)}>
        <div class="coworker__person">
          <mgt-person
            .personDetails=${person}
            .line2Property=${'jobTitle'}
            .fetchImage=${true}
            .view=${ViewType.twolines}
          ></mgt-person>
        </div>
      </div>
    `;
  }

  /**
   * Render a coworker org member
   *
   * @protected
   * @param {User} person
   * @returns {TemplateResult}
   * @memberof MgtPersonCardOrganization
   */
  protected renderCoworkers(): TemplateResult {
    const { people } = this._state;
    if (!people || !people.length) {
      return null;
    }

    const subtitle =
      this._me.id === this._state.person.id
        ? this.strings.youWorkWithSubSectionTitle
        : `${this._state.person.givenName} ${this.strings.userWorksWithSubSectionTitle}`;

    return html`
      <div class="divider"></div>
      <div class="subtitle" tabindex="0">${subtitle}</div>
      <div>
        ${people.slice(0, 6).map(person => this.renderCoworker(person))}
      </div>
    `;
  }
}
