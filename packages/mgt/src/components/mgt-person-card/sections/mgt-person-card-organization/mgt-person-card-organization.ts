/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, TemplateResult } from 'lit-element';
import { Providers } from '../../../../Providers';
import { ProviderState } from '../../../../providers/IProvider';
import { BasePersonCardSection } from '../BasePersonCardSection';
import { getCoworkers, getManagers, IOrgMember } from './graph.organization';
import { styles } from './mgt-person-card-organization-css';

/**
 * foo
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
   * foo
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
   * foo
   *
   * @returns {TemplateResult}
   * @memberof MgtPersonCardOrganization
   */
  public renderIcon(): TemplateResult {
    return html`
      <svg xmlns="http://www.w3.org/2000/svg">
        <path
          fill-rule="evenodd"
          clip-rule="evenodd"
          d="M13 4H8V7H13V4ZM8 3C7.44772 3 7 3.44772 7 4V7C7 7.55228 7.44772 8 8 8H10V9H7.5C6.67157 9 6 9.67157 6 10.5V11H4C3.44772 11 3 11.4477 3 12V15C3 15.5523 3.44772 16 4 16H9C9.55228 16 10 15.5523 10 15V12C10 11.4477 9.55228 11 9 11H7V10.5C7 10.2239 7.22386 10 7.5 10H13.5C13.7761 10 14 10.2239 14 10.5V11H12C11.4477 11 11 11.4477 11 12V15C11 15.5523 11.4477 16 12 16H17C17.5523 16 18 15.5523 18 15V12C18 11.4477 17.5523 11 17 11H15V10.5C15 9.67157 14.3284 9 13.5 9H11V8H13C13.5523 8 14 7.55228 14 7V4C14 3.44772 13.5523 3 13 3H8ZM9 12H4L4 15H9V12ZM12 12H17V15H12V12Z"
        />
      </svg>
    `;
  }

  /**
   * foo
   *
   * @protected
   * @memberof MgtPersonCardOrganization
   */
  public clearState(): void {
    this._managers = [];
    this._coworkers = [];
  }

  /**
   * foo
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
   * foo
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
          <svg width="8" height="13" viewBox="0 0 8 13" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M1 12L6.5 6.5L1 1" stroke="#B8B8B8" stroke-width="2" />
          </svg>
        </div>
      </div>
      <div class="org-member__separator"></div>
    `;
  }

  /**
   * foo
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
   * foo
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
