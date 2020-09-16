/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, TemplateResult } from 'lit-element';
import { Providers, ProviderState } from '@microsoft/mgt-element';
import { BasePersonCardSection } from '../BasePersonCardSection';
import { getCoworkers, getManagers, IOrgMember } from './graph.organization';
import { styles } from './mgt-person-card-organization-css';
import { getSvg, SvgIcon } from '../../../../utils/SvgHelper';

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
   * The name for display in the overview section.
   *
   * @readonly
   * @type {string}
   * @memberof MgtPersonCardOrganization
   */
  public get displayName(): string {
    return 'Reports to';
  }

  private _managers: IOrgMember[];
  private _coworkers: IOrgMember[];

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
  public clearState(): void {
    this._managers = [];
    this._coworkers = [];
  }

  /**
   * Render the compact view
   *
   * @returns {TemplateResult}
   * @memberof MgtPersonCardOrganization
   */
  protected renderCompactView(): TemplateResult {
    let contentTemplate: TemplateResult;

    if (this.isLoadingState) {
      contentTemplate = this.renderLoading();
    } else if (!this._managers || !this._managers.length) {
      contentTemplate = this.renderNoData();
    } else {
      const reportsTo = this._managers[0];
      contentTemplate = this.renderCoworker(reportsTo);
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

    if (this.isLoadingState) {
      contentTemplate = this.renderLoading();
    } else if ((!this._managers || !this._managers.length) && (!this._coworkers || !this._coworkers.length)) {
      contentTemplate = this.renderNoData();
    } else {
      const managers = new Array(...this._managers);
      const managerTemplates = managers ? managers.reverse().map(manager => this.renderManager(manager)) : [];
      const targetMemberTemplate = this.personDetails ? this.renderTargetMember() : null;
      const coworkersTemplate =
        this._coworkers && this._coworkers.length
          ? html`
              <div class="divider"></div>
              <div class="subtitle">You work with</div>
              <div>
                ${this._coworkers.slice(0, 6).map(coworker => this.renderCoworker(coworker))}
              </div>
            `
          : [];

      contentTemplate = html`
        ${managerTemplates} ${targetMemberTemplate} ${coworkersTemplate}
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
   * @param {IOrgMember} orgMember
   * @returns {TemplateResult}
   * @memberof MgtPersonCardOrganization
   */
  protected renderManager(orgMember: IOrgMember): TemplateResult {
    return html`
      <div class="org-member" @click=${() => this.navigateCard(orgMember)}>
        <div class="org-member__image">
          <mgt-person .userId=${orgMember.id} avatar-size="large"></mgt-person>
        </div>
        <div class="org-member__details">
          <div class="org-member__name">${orgMember.displayName}</div>
          <div class="org-member__title">${orgMember.title}</div>
          <div class="org-member__department">${orgMember.department}</div>
        </div>
        <div class="org-member__more">
          ${getSvg(SvgIcon.ExpandRight)}
        </div>
      </div>
      <div class="org-member__separator"></div>
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
    return html`
      <div class="org-member org-member--target">
        <div class="org-member__image">
          <mgt-person .personDetails=${this.personDetails} avatar-size="large"></mgt-person>
        </div>
        <div class="org-member__details">
          <div class="org-member__name">${this.personDetails.displayName}</div>
          <div class="org-member__title">${this.personDetails.jobTitle}</div>
          <div class="org-member__department">${this.personDetails.department}</div>
        </div>
      </div>
    `;
  }

  /**
   * Render a coworker org member
   *
   * @protected
   * @param {IOrgMember} coworker
   * @returns {TemplateResult}
   * @memberof MgtPersonCardOrganization
   */
  protected renderCoworker(coworker: IOrgMember): TemplateResult {
    return html`
      <div class="coworker" @click=${() => this.navigateCard(coworker)}>
        <div class="coworker__image">
          <mgt-person .userId=${coworker.id} avatar-size="large"></mgt-person>
        </div>
        <div class="coworker__details">
          <div class="coworker__name">${coworker.displayName}</div>
          <div class="coworker__title">${coworker.title}</div>
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
    const provider = Providers.globalProvider;

    // check if user is signed in
    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    if (!this.personDetails) {
      return;
    }

    const graph = provider.graph.forComponent(this);
    const userId = this.personDetails.id;
    this._managers = await getManagers(graph, userId);
    this._coworkers = await getCoworkers(graph, userId);

    this.requestUpdate();
  }
}
