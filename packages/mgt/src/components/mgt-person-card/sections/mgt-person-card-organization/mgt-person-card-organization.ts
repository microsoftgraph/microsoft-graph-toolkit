/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, TemplateResult } from 'lit-element';
import { Providers, ProviderState } from '@microsoft/mgt-element';
import { BasePersonCardSection } from '../BasePersonCardSection';
import { styles } from './mgt-person-card-organization-css';
import { getSvg, SvgIcon } from '../../../../utils/SvgHelper';
import { MgtPersonCardState } from '../../mgt-person-card.graph';
import { User } from '@microsoft/microsoft-graph-types';
import { PersonViewType } from '../../../mgt-person/mgt-person';

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

  private _state: MgtPersonCardState;

  constructor(state: MgtPersonCardState) {
    super();
    this._state = state;
  }

  /**
   * The name for display in the overview section.
   *
   * @readonly
   * @type {string}
   * @memberof MgtPersonCardOrganization
   */
  public get displayName(): string {
    return 'Reports to';
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
   * Reset any state in the section
   *
   * @protected
   * @memberof MgtPersonCardOrganization
   */
  public clearState(): void {}

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

    const { person, directReports, people } = this._state;

    if (!person) {
      return null;
    } else if (person && person.manager) {
      contentTemplate = this.renderCoworker(person.manager);
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
      const managers = [];
      let currentManager = person;
      while (currentManager.manager) {
        managers.push(currentManager.manager);
        currentManager = currentManager.manager;
      }

      const managerTemplates = managers ? managers.reverse().map(manager => this.renderManager(manager)) : [];
      const targetMemberTemplate = this.renderTargetMember();
      const directReportsTemplate = this.renderDirectReports();
      const coworkersTemplate =
        people && people.length
          ? html`
              <div class="divider"></div>
              <div class="subtitle">You work with</div>
              <div>
                ${people.slice(0, 6).map(coworker => this.renderCoworker(coworker))}
              </div>
            `
          : [];

      contentTemplate = html`
        ${managerTemplates} ${targetMemberTemplate} ${directReportsTemplate} ${coworkersTemplate}
      `;
    }

    return html`
      <div class="root">
        <div class="title">Organization</div>
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
      <div class="org-member" @click=${() => this.navigateCard(person)}>
        <div class="org-member__image">
          <mgt-person .userId=${person.id} avatar-size="large"></mgt-person>
        </div>
        <div class="org-member__details">
          <div class="org-member__name">${person.displayName}</div>
          <div class="org-member__title">${person.jobTitle}</div>
          <div class="org-member__department">${person.department}</div>
        </div>
        <div class="org-member__more">
          ${getSvg(SvgIcon.ExpandRight)}
        </div>
      </div>
      <div class="org-member__separator"></div>
    `;
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
    if (!directReports) {
      return null;
    }

    return html`
      <div class="org-member__separator"></div>
      <div>
        ${directReports.map(
          person => html`
            <div class="org-member org-member--direct-report" @click=${() => this.navigateCard(person)}>
              <div class="org-member__image">
                <mgt-person
                  .personDetails=${person}
                  .line2Property=${'jobTitle'}
                  .line3Property=${'department'}
                  .fetchImage=${true}
                  .view=${PersonViewType.twolines}
                ></mgt-person>
              </div>
              <div class="org-member__more">
                ${getSvg(SvgIcon.ExpandRight)}
              </div>
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
  protected renderTargetMember(): TemplateResult {
    const { person } = this._state;
    // TODO - update to three lines
    return html`
      <div class="org-member org-member--target">
        <div class="org-member__image">
          <mgt-person
            .personDetails=${person}
            .line2Property=${'jobTitle'}
            .line3Property=${'department'}
            .fetchImage=${true}
            .view=${PersonViewType.twolines}
          ></mgt-person>
        </div>
      </div>
    `;

    // <!-- <div class="org-member__details">
    // <div class="org-member__name">${this.personDetails.displayName}</div>
    // <div class="org-member__title">${this.personDetails.jobTitle}</div>
    // <div class="org-member__department">${this.personDetails.department}</div>
    // </div> -->
  }

  /**
   * Render a coworker org member
   *
   * @protected
   * @param {User} coworker
   * @returns {TemplateResult}
   * @memberof MgtPersonCardOrganization
   */
  protected renderCoworker(coworker: User): TemplateResult {
    return html`
      <div class="coworker" @click=${() => this.navigateCard(coworker)}>
        <div class="coworker__image">
          <mgt-person .personDetails=${coworker} avatar-size="large" fetch-image></mgt-person>
        </div>
        <div class="coworker__details">
          <div class="coworker__name">${coworker.displayName}</div>
          <div class="coworker__title">${coworker.jobTitle}</div>
        </div>
      </div>
    `;
  }

  /**
   * load state into the component
   *
   * @protected
   * @returns {Promise<void>}
   * @memberof MgtPersonCardProfile
   */
  protected async loadState(): Promise<void> {
    // const provider = Providers.globalProvider;
    // // check if user is signed in
    // if (!provider || provider.state !== ProviderState.SignedIn) {
    //   return;
    // }
    // if (!this.personDetails) {
    //   return;
    // }
    // const graph = provider.graph.forComponent(this);
    // const userId = this.personDetails.id;
    // this._managers = await getManagers(graph, userId);
    // this._coworkers = await getCoworkers(graph, userId);
    // this.requestUpdate();
  }
}
