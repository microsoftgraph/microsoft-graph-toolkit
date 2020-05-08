/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, TemplateResult } from 'lit-element';
import { Providers } from '../../../../../Providers';
import { ProviderState } from '../../../../../providers/IProvider';
import { BetaGraph } from '../../../../BetaGraph';
import { BasePersonCardSection } from '../BasePersonCardSection';
import { styles } from './mgt-person-card-organization-css';

/**
 * Defines the data required to render an org member.
 *
 * @interface IOrgMember
 */
interface IOrgMember {
  id: string;
  image: string;
  displayName: string;
  title: string;
  department?: string;
}

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

  private _orgMembers: IOrgMember[];
  private _coworkers: IOrgMember[];

  /**
   * foo
   *
   * @returns {TemplateResult}
   * @memberof MgtPersonCardOrganization
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
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    if (this.isCompact) {
      if (!this._orgMembers) {
        return html`
          <div>show shimmer here</div>
        `;
      }

      const reportsTo = this._orgMembers[0];
      return html`
        <div class="root compact">
          ${this.renderCoworker(reportsTo)}
        </div>
      `;
    }

    const orgMemberTemplates = this._orgMembers
      ? this._orgMembers.map(orgMember => this.renderOrgMember(orgMember))
      : [];

    const targetMemberTemplate = this.personDetails ? this.renderTargetMember() : null;
    const coworkerTemplates = this._coworkers ? this._coworkers.map(coworker => this.renderCoworker(coworker)) : [];

    return html`
      <div class="root">
        <div class="title">Organization</div>
        <div>
          ${orgMemberTemplates} ${targetMemberTemplate}
        </div>
        <div class="divider"></div>
        <div class="subtitle">You work with</div>
        <div>
          ${coworkerTemplates}
        </div>
      </div>
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
  protected renderOrgMember(orgMember: IOrgMember): TemplateResult {
    return html`
      <div class="org-member">
        <div class="org-member__image">${orgMember.image}</div>
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
      <div class="org-member">
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
      <div class="coworker">
        <div class="coworker__image"></div>
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
    const betaGraph = BetaGraph.fromGraph(graph);

    // TODO: Get real data

    this._orgMembers = [];
    this._coworkers = [];

    this.injectDummyData();

    this.requestUpdate();
  }

  private injectDummyData() {
    const orgMembers: IOrgMember[] = [
      {
        department: 'EPIC Mgmt RnD',
        displayName: 'Jessica Smith',
        id: '',
        image: '',
        title: 'SR PM MANAGER'
      },
      {
        department: 'EPIC Mgmt RnD',
        displayName: 'Jessica Smith',
        id: '',
        image: '',
        title: 'SR PM MANAGER'
      },
      {
        department: 'EPIC Mgmt RnD',
        displayName: 'Jessica Smith',
        id: '',
        image: '',
        title: 'SR PM MANAGER'
      },
      {
        department: 'EPIC Mgmt RnD',
        displayName: 'Jessica Smith',
        id: '',
        image: '',
        title: 'SR PM MANAGER'
      },
      {
        department: 'EPIC Mgmt RnD',
        displayName: 'Jessica Smith',
        id: '',
        image: '',
        title: 'SR PM MANAGER'
      }
    ];

    const coworkers: IOrgMember[] = [
      {
        displayName: 'Jessica Smith',
        id: '',
        image: '',
        title: 'SR PM MANAGER'
      },
      {
        displayName: 'Jessica Smith',
        id: '',
        image: '',
        title: 'SR PM MANAGER'
      },
      {
        displayName: 'Jessica Smith',
        id: '',
        image: '',
        title: 'SR PM MANAGER'
      },
      {
        displayName: 'Jessica Smith',
        id: '',
        image: '',
        title: 'SR PM MANAGER'
      },
      {
        displayName: 'Jessica Smith',
        id: '',
        image: '',
        title: 'SR PM MANAGER'
      }
    ];

    this._orgMembers = orgMembers;
    this._coworkers = coworkers;
  }
}
