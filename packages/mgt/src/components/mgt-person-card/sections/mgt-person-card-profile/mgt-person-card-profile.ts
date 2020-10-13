/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, TemplateResult } from 'lit-element';
import { BetaGraph } from '../../../../BetaGraph';
import { BasePersonCardSection } from '../BasePersonCardSection';
import { getProfile, IPersonAnniversary, IPersonInterest, IProfile } from './graph.profile';
import { styles } from './mgt-person-card-profile-css';
import { ProviderState, Providers } from '@microsoft/mgt-element';
import { getSvg, SvgIcon } from '../../../../utils/SvgHelper';

/**
 * The user profile subsection of the person card
 *
 * @export
 * @class MgtPersonCardProfile
 * @extends {MgtTemplatedComponent}
 */
@customElement('mgt-person-card-profile')
export class MgtPersonCardProfile extends BasePersonCardSection {
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
   * @memberof MgtPersonCardProfile
   */
  public get displayName(): string {
    return 'Skills & Experience';
  }

  /**
   * The user's profile metadata
   *
   * @protected
   * @type {IProfile}
   * @memberof MgtPersonCardProfile
   */
  protected get profile(): IProfile {
    return this._profile;
  }
  protected set profile(value: IProfile) {
    if (value === this._profile) {
      return;
    }

    this._profile = value;
    this._birthdayAnniversary =
      value && value.anniversaries ? value.anniversaries.find(this.isBirthdayAnniversary) : null;
    this._personalInterests = value && value.interests ? value.interests.filter(this.isPersonalInterest) : null;
    this._professionalInterests = value && value.interests ? value.interests.filter(this.isProfessionalInterest) : null;
  }

  private _profile: IProfile;
  private _personalInterests: IPersonInterest[];
  private _professionalInterests: IPersonInterest[];
  private _birthdayAnniversary: IPersonAnniversary;

  constructor() {
    super();

    this.profile = null;
  }

  /**
   * Render the icon for display in the navigation ribbon.
   *
   * @returns {TemplateResult}
   * @memberof MgtPersonCardProfile
   */
  public renderIcon(): TemplateResult {
    return getSvg(SvgIcon.Profile);
  }

  /**
   * Reset any state in the section
   *
   * @protected
   * @memberof MgtPersonCardProfile
   */
  public clearState(): void {
    this.profile = null;
  }

  /**
   * Render the compact view
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCardProfile
   */
  protected renderCompactView(): TemplateResult {
    let contentTemplate: TemplateResult;

    if (this.isLoadingState) {
      contentTemplate = this.renderLoading();
    } else if (!this._profile) {
      contentTemplate = this.renderNoData();
    } else {
      this.initPostRenderOperations();
      contentTemplate = html`
        ${this.renderSkills()} ${this.renderProfessionalInterests()} ${this.renderBirthday()}
      `;
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
   * @returns
   * @memberof MgtPersonCardProfile
   */
  protected renderFullView() {
    let contentTemplate: TemplateResult;

    if (this.isLoadingState) {
      contentTemplate = this.renderLoading();
    } else if (!this._profile) {
      contentTemplate = this.renderNoData();
    } else {
      this.initPostRenderOperations();
      contentTemplate = html`
        ${this.renderLanguages()} ${this.renderSkills()} ${this.renderWorkExperience()} ${this.renderEducation()}
        ${this.renderProfessionalInterests()} ${this.renderPersonalInterests()} ${this.renderBirthday()}
      `;
    }

    return html`
      <div class="root">
        <div class="title">About</div>
        ${contentTemplate}
      </div>
    `;
  }

  /**
   * Render the user's known languages
   *
   * @protected
   * @returns
   * @memberof MgtPersonCardProfile
   */
  protected renderLanguages(): TemplateResult {
    let contentTemplate: TemplateResult;

    if (this._profile && this._profile.languages) {
      const languageItems: TemplateResult[] = [];

      for (const language of this._profile.languages) {
        let proficiency = null;
        if (language.proficiency && language.proficiency.length) {
          proficiency = html`
            <span class="language__proficiency">
              &nbsp;(${language.proficiency})
            </span>
          `;
        }

        languageItems.push(html`
          <div class="token-list__item language">
            <span class="language__title">${language.displayName}</span>
            ${proficiency}
          </div>
        `);
      }

      contentTemplate = html`
        <div class="token-list">
          ${languageItems}
        </div>
      `;
    } else {
      contentTemplate = html`
        <div>None</div>
      `;
    }

    const languagesTemplate = html`
      <section>
        <div class="section__title">Languages</div>
        <div class="section__content">
          ${contentTemplate}
        </div>
      </section>
    `;

    return languagesTemplate;
  }

  /**
   * Render the user's skills
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCardProfile
   */
  protected renderSkills(): TemplateResult {
    let contentTemplate: TemplateResult;

    if (this._profile && this._profile.skills) {
      const skillItems: TemplateResult[] = [];
      for (const skill of this._profile.skills) {
        skillItems.push(html`
          <div class="token-list__item skill">
            ${skill.displayName}
          </div>
        `);
      }

      contentTemplate = html`
        <div class="token-list">
          ${skillItems}
        </div>
      `;
    } else {
      contentTemplate = html`
        <div>None</div>
      `;
    }

    return html`
      <section>
        <div class="section__title">Skills</div>
        <div class="section__content">
          ${contentTemplate}
        </div>
      </section>
    `;
  }

  /**
   * Render the user's work experience timeline
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCardProfile
   */
  protected renderWorkExperience(): TemplateResult {
    let contentTemplate: TemplateResult;

    if (this._profile && this._profile.positions) {
      const positionItems: TemplateResult[] = [];
      for (const position of this._profile.positions) {
        positionItems.push(html`
          <div class="data-list__item work-position">
            <div class="data-list__item__title">${position.detail.jobTitle}</div>
            <div class="data-list__item__content flex-rows">
              <div>
                <div class="work-position__company">${position.detail.company.displayName}</div>
                <div class="work-position__location">
                  ${position.detail.company.address.city}, ${position.detail.company.address.state}
                </div>
              </div>
              <div class="date-range">
                <div class="work-position__date-range">${this.getDisplayDateRange(position.detail)}</div>
                <div class="date-range__circle"></div>
              </div>
            </div>
          </div>
        `);
      }

      contentTemplate = html`
        <div class="data-list" data-if="positions && positions.length">
          ${positionItems}
        </div>
      `;
    } else {
      contentTemplate = html`
        <div>None</div>
      `;
    }

    return html`
      <section>
        <div class="section__title">Work Experience</div>
        <div class="section__content">
          ${contentTemplate}
        </div>
      </section>
    `;
  }

  /**
   * Render the user's education timeline
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCardProfile
   */
  protected renderEducation(): TemplateResult {
    let contentTemplate: TemplateResult;

    if (this._profile && this._profile.positions) {
      const positionItems: TemplateResult[] = [];
      for (const educationalActivity of this._profile.educationalActivities) {
        positionItems.push(html`
          <div class="data-list__item educational-activity">
            <div class="data-list__item__title">${educationalActivity.institution.displayName}</div>
            <div class="data-list__item__content flex-rows">
              <div>
                <div class="educational-activity__degree">${educationalActivity.program.displayName}</div>
              </div>
              <div class="date-range">
                <div class="educational-activity__date-range">
                  ${this.getDisplayDateRange(educationalActivity)}
                </div>
                <div class="date-range__circle"></div>
              </div>
            </div>
          </div>
        `);
      }

      contentTemplate = html`
        <div class="data-list">
          ${positionItems}
        </div>
      `;
    } else {
      contentTemplate = html`
        <div>None</div>
      `;
    }

    return html`
      <section>
        <div class="section__title">Education</div>
        <div class="section__content">
          ${contentTemplate}
        </div>
      </section>
    `;
  }

  /**
   * Render the user's professional interests
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCardProfile
   */
  protected renderProfessionalInterests(): TemplateResult {
    let contentTemplate: TemplateResult;

    if (this._professionalInterests && this._professionalInterests.length) {
      const interestItems: TemplateResult[] = [];
      for (const interest of this._professionalInterests) {
        interestItems.push(html`
          <div class="token-list__item interest interest--professional">
            ${interest.displayName}
          </div>
        `);
      }

      contentTemplate = html`
        <div class="token-list">
          ${interestItems}
        </div>
      `;
    } else {
      contentTemplate = html`
        <div>None</div>
      `;
    }

    return html`
      <section>
        <div class="section__title">Professional Interests</div>
        <div class="section__content">
          ${contentTemplate}
        </div>
      </section>
    `;
  }

  /**
   * Render the user's personal interests
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCardProfile
   */
  protected renderPersonalInterests(): TemplateResult {
    let contentTemplate: TemplateResult;

    if (this._personalInterests && this._personalInterests.length) {
      const interestItems: TemplateResult[] = [];
      for (const interest of this._personalInterests) {
        interestItems.push(html`
          <div class="token-list__item interest interest--personal">
            ${interest.displayName}
          </div>
        `);
      }

      contentTemplate = html`
        <div class="token-list">
          ${interestItems}
        </div>
      `;
    } else {
      contentTemplate = html`
        <div>None</div>
      `;
    }

    return html`
      <section>
        <div class="section__title">Personal Interests</div>
        <div class="section__content">
          ${contentTemplate}
        </div>
      </section>
    `;
  }

  /**
   * Render the user's birthday
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCardProfile
   */
  protected renderBirthday(): TemplateResult {
    let contentTemplate: TemplateResult;

    if (this._birthdayAnniversary) {
      contentTemplate = html`
        <div class="birthday">
          <div class="birthday__icon">
            ${getSvg(SvgIcon.Birthday)}
          </div>
          <div class="birthday__date">
            ${this.getDisplayDate(this._birthdayAnniversary.date)}
          </div>
        </div>
      `;
    } else {
      contentTemplate = html`
        <div>Unknown</div>
      `;
    }

    return html`
      <section>
        <div class="section__title">Birthday</div>
        <div class="section__content">
          ${contentTemplate}
        </div>
      </section>
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

    const userId = this.personDetails.id;
    const profile = await getProfile(betaGraph, userId);

    this.profile = profile;
    this.requestUpdate();
  }

  private isPersonalInterest(interest: IPersonInterest): boolean {
    return interest.categories && interest.categories.includes('personal');
  }

  private isProfessionalInterest(interest: IPersonInterest): boolean {
    return interest.categories && interest.categories.includes('professional');
  }

  private isBirthdayAnniversary(anniversary: IPersonAnniversary): boolean {
    return anniversary.type === 'birthday';
  }

  private getDisplayDate(date: Date): string {
    return date.toLocaleString('default', {
      day: 'numeric',
      month: 'long'
    });
  }

  // tslint:disable-next-line: completed-docs
  private getDisplayDateRange(event: { startMonthYear: Date; endMonthYear: Date }): string {
    const start = new Date(event.startMonthYear).getFullYear();
    if (start === 0) {
      return null;
    }

    const end = event.endMonthYear ? new Date(event.endMonthYear).getFullYear() : 'Current';
    return `${start} — ${end}`;
  }

  private initPostRenderOperations(): void {
    setTimeout(() => {
      try {
        const sections = this.shadowRoot.querySelectorAll('section');
        sections.forEach(section => {
          // Perform post render operations per section
          this.handleTokenOverflow(section);
          this.drawSectionTimeline(section);
        });
      } catch {
        // An exception may occur if the component is suddenly removed during post render operations.
      }
    }, 0);
  }

  private handleTokenOverflow(section: HTMLElement): void {
    const tokenLists = section.querySelectorAll('.token-list');
    if (!tokenLists || !tokenLists.length) {
      return;
    }

    for (const tokenList of Array.from(tokenLists)) {
      const items = tokenList.querySelectorAll('.token-list__item');
      if (!items || !items.length) {
        continue;
      }

      let overflowItems: Element[] = null;
      let itemRect = items[0].getBoundingClientRect();
      const tokenListRect = tokenList.getBoundingClientRect();
      const maxtop = itemRect.height * 2 + tokenListRect.top;

      // Use (items.length - 1) to prevent [+1 more] from appearing.
      for (let i = 0; i < items.length - 1; i++) {
        itemRect = items[i].getBoundingClientRect();
        if (itemRect.top > maxtop) {
          overflowItems = Array.from(items).slice(i, items.length);
          break;
        }
      }

      if (overflowItems) {
        overflowItems.forEach(i => i.classList.add('overflow'));

        const overflowToken = document.createElement('div');
        overflowToken.classList.add('token-list__item');
        overflowToken.innerText = `+ ${overflowItems.length} more`;
        overflowToken.addEventListener('click', (e: MouseEvent) => {
          // On click, remove [+n more] token and reveal the hidden overflow tokens.
          overflowToken.remove();
          overflowItems.forEach(i => i.classList.remove('overflow'));
        });
        tokenList.appendChild(overflowToken);
      }
    }
  }

  private drawSectionTimeline(section) {
    const circles = section.querySelectorAll('.date-range .date-range__circle');
    if (!circles || circles.length <= 1) {
      return;
    }

    for (let i = 0; i < circles.length - 1; i++) {
      const currentCircle = circles[i];
      const nextCircle = circles[i + 1];

      const line = document.createElement('div');
      line.classList.add('date-range__line');
      section.appendChild(line);

      const lineRect = line.getBoundingClientRect();
      const lineParentRect = line.offsetParent.getBoundingClientRect();
      const topCircleRect = currentCircle.getBoundingClientRect();
      const lastCircleRect = nextCircle.getBoundingClientRect();
      const top = topCircleRect.bottom - lineParentRect.top;
      const height = lastCircleRect.top - lineParentRect.top - top;
      const left = topCircleRect.left - lineParentRect.left + (topCircleRect.width / 2 - lineRect.width / 2);

      line.style.top = `${top}px`;
      line.style.height = `${height}px`;
      line.style.left = `${left}px`;
    }
  }
}
