/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { EducationalActivity, PersonAnnualEvent, PersonInterest, Profile } from '@microsoft/microsoft-graph-types-beta';
import { customElement, html, TemplateResult } from 'lit-element';
import { BasePersonCardSection } from '../BasePersonCardSection';
import { getSvg, SvgIcon } from '../../../../utils/SvgHelper';
import { styles } from './mgt-person-card-profile-css';
import { strings } from './strings';

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

  protected get strings() {
    return strings;
  }

  /**
   * The name for display in the overview section.
   *
   * @readonly
   * @type {string}
   * @memberof MgtPersonCardProfile
   */
  public get displayName(): string {
    return this.strings.SkillsAndExperienceSectionTitle;
  }

  /**
   * Returns true if the profile contains data
   * that can be rendered
   *
   * @readonly
   * @type {boolean}
   * @memberof MgtPersonCardProfile
   */
  public get hasData(): boolean {
    if (!this.profile) {
      return false;
    }

    const { languages, skills, positions, educationalActivities } = this.profile;

    return (
      [
        this._birthdayAnniversary,
        this._personalInterests && this._personalInterests.length,
        this._professionalInterests && this._professionalInterests.length,
        languages && languages.length,
        skills && skills.length,
        positions && positions.length,
        educationalActivities && educationalActivities.length
      ].filter(v => !!v).length > 0
    );
  }

  /**
   * The user's profile metadata
   *
   * @protected
   * @type {IProfile}
   * @memberof MgtPersonCardProfile
   */
  protected get profile(): Profile {
    return this._profile;
  }
  protected set profile(value: Profile) {
    if (value === this._profile) {
      return;
    }

    this._profile = value;
    this._birthdayAnniversary =
      value && value.anniversaries ? value.anniversaries.find(this.isBirthdayAnniversary) : null;
    this._personalInterests = value && value.interests ? value.interests.filter(this.isPersonalInterest) : null;
    this._professionalInterests = value && value.interests ? value.interests.filter(this.isProfessionalInterest) : null;
  }

  private _profile: Profile;
  private _personalInterests: PersonInterest[];
  private _professionalInterests: PersonInterest[];
  private _birthdayAnniversary: PersonAnnualEvent;

  constructor(profile: Profile) {
    super();

    this.profile = profile;
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
    super.clearState();
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
    return html`
      <div class="root compact" dir=${this.direction}>
        ${this.renderSubSections().slice(0, 2)}
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
    this.initPostRenderOperations();

    return html`
      <div class="root" dir=${this.direction}>
        <div class="title">${this.strings.AboutCompactSectionTitle}</div>
        ${this.renderSubSections()}
      </div>
    `;
  }

  /**
   * Renders all subSections of the profile
   * Defines order of how they render
   *
   * @protected
   * @return {*}
   * @memberof MgtPersonCardProfile
   */
  protected renderSubSections() {
    const subSections = [
      this.renderSkills(),
      this.renderBirthday(),
      this.renderLanguages(),
      this.renderWorkExperience(),
      this.renderEducation(),
      this.renderProfessionalInterests(),
      this.renderPersonalInterests()
    ];

    return subSections.filter(s => !!s);
  }

  /**
   * Render the user's known languages
   *
   * @protected
   * @returns
   * @memberof MgtPersonCardProfile
   */
  protected renderLanguages(): TemplateResult {
    const { languages } = this._profile;
    if (!(languages && languages.length)) {
      return null;
    }

    const languageItems: TemplateResult[] = [];
    for (const language of languages) {
      let proficiency = null;
      if (language.proficiency && language.proficiency.length) {
        proficiency = html`
          <span class="language__proficiency" tabindex="0">
            &nbsp;(${language.proficiency})
          </span>
        `;
      }

      languageItems.push(html`
        <div class="token-list__item language">
          <span class="language__title" tabindex="0">${language.displayName}</span>
          ${proficiency}
        </div>
      `);
    }

    const languageTitle = languageItems.length ? this.strings.LanguagesSubSectionTitle : '';

    return html`
      <section>
        <div class="section__title" tabindex="0">${languageTitle}</div>
        <div class="section__content">
          <div class="token-list">
            ${languageItems}
          </div>
        </div>
      </section>
    `;
  }

  /**
   * Render the user's skills
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCardProfile
   */
  protected renderSkills(): TemplateResult {
    const { skills } = this._profile;

    if (!(skills && skills.length)) {
      return null;
    }

    const skillItems: TemplateResult[] = [];
    for (const skill of skills) {
      skillItems.push(html`
        <div class="token-list__item skill" tabindex="0">
          ${skill.displayName}
        </div>
      `);
    }

    const skillsTitle = skillItems.length ? this.strings.SkillsSubSectionTitle : '';

    return html`
      <section>
        <div class="section__title" tabindex="0">${skillsTitle}</div>
        <div class="section__content">
          <div class="token-list">
            ${skillItems}
          </div>
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
    const { positions } = this._profile;

    if (!(positions && positions.length)) {
      return null;
    }

    const positionItems: TemplateResult[] = [];
    for (const position of this._profile.positions) {
      if (position.detail.description || position.detail.jobTitle !== '') {
        positionItems.push(html`
          <div class="data-list__item work-position">
            <div class="data-list__item__header">
              <div class="data-list__item__title" tabindex="0">${position.detail?.jobTitle}</div>
              <div class="data-list__item__date-range" tabindex="0">
                ${this.getDisplayDateRange(position.detail)}
              </div>
            </div>
            <div class="data-list__item__content">
              <div class="work-position__company" tabindex="0">
                ${position?.detail?.company?.displayName}
              </div>
              <div class="work-position__location" tabindex="0">
                ${position?.detail?.company?.address?.city}, ${position?.detail?.company?.address?.state}
              </div>
            </div>
          </div>
        `);
      }
    }
    const workExperienceTitle = positionItems.length ? this.strings.WorkExperienceSubSectionTitle : '';

    return html`
      <section>
        <div class="section__title" tabindex="0">${workExperienceTitle}</div>
        <div class="section__content">
          <div class="data-list">
            ${positionItems}
          </div>
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
    const { educationalActivities } = this._profile;

    if (!(educationalActivities && educationalActivities.length)) {
      return null;
    }

    const positionItems: TemplateResult[] = [];
    for (const educationalActivity of educationalActivities) {
      positionItems.push(html`
        <div class="data-list__item educational-activity">
          <div class="data-list__item__header">
            <div class="data-list__item__title" tabindex="0">${educationalActivity.institution.displayName}</div>
            <div class="data-list__item__date-range" tabindex="0">
              ${this.getDisplayDateRange(educationalActivity)}
            </div>
          </div>
          <div class="data-list__item__content">
            <div class="educational-activity__degree" tabindex="0">
              ${educationalActivity.program.displayName || 'Bachelors Degree'}
            </div>
          </div>
        </div>
      `);
    }

    const educationTitle = positionItems.length ? this.strings.EducationSubSectionTitle : '';

    return html`
      <section>
        <div class="section__title" tabindex="0">${educationTitle}</div>
        <div class="section__content">
          <div class="data-list">
            ${positionItems}
          </div>
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
    if (!this._professionalInterests || !this._professionalInterests.length) {
      return null;
    }

    const interestItems: TemplateResult[] = [];
    for (const interest of this._professionalInterests) {
      interestItems.push(html`
        <div class="token-list__item interest interest--professional" tabindex="0">
          ${interest.displayName}
        </div>
      `);
    }

    const professionalInterests = interestItems.length ? this.strings.professionalInterestsSubSectionTitle : '';

    return html`
      <section>
        <div class="section__title" tabindex="0">${professionalInterests}</div>
        <div class="section__content">
          <div class="token-list">
            ${interestItems}
          </div>
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
    if (!this._personalInterests || !this._personalInterests.length) {
      return null;
    }

    const interestItems: TemplateResult[] = [];
    for (const interest of this._personalInterests) {
      interestItems.push(html`
        <div class="token-list__item interest interest--personal" tabindex="0">
          ${interest.displayName}
        </div>
      `);
    }

    const personalInterests = interestItems.length ? this.strings.personalInterestsSubSectionTitle : '';

    return html`
      <section>
        <div class="section__title" tabindex="0">${personalInterests}</div>
        <div class="section__content">
          <div class="token-list">
            ${interestItems}
          </div>
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
    if (!this._birthdayAnniversary || !this._birthdayAnniversary.date) {
      return null;
    }

    return html`
      <section>
        <div class="section__title" tabindex="0">Birthday</div>
        <div class="section__content">
          <div class="birthday">
            <div class="birthday__icon">
              ${getSvg(SvgIcon.Birthday)}
            </div>
            <div class="birthday__date" tabindex="0">
              ${this.getDisplayDate(new Date(this._birthdayAnniversary.date))}
            </div>
          </div>
        </div>
      </section>
    `;
  }

  private isPersonalInterest(interest: PersonInterest): boolean {
    return interest.categories && interest.categories.includes('personal');
  }

  private isProfessionalInterest(interest: PersonInterest): boolean {
    return interest.categories && interest.categories.includes('professional');
  }

  private isBirthdayAnniversary(anniversary: PersonAnnualEvent): boolean {
    return anniversary.type === 'birthday';
  }

  private getDisplayDate(date: Date): string {
    return date.toLocaleString('default', {
      day: 'numeric',
      month: 'long'
    });
  }

  // tslint:disable-next-line: completed-docs
  private getDisplayDateRange(event: EducationalActivity): string {
    const start = new Date(event.startMonthYear).getFullYear();
    if (start === 0) {
      return null;
    }

    const end = event.endMonthYear ? new Date(event.endMonthYear).getFullYear() : this.strings.currentYearSubtitle;
    return `${start} — ${end}`;
  }

  private initPostRenderOperations(): void {
    setTimeout(() => {
      try {
        const sections = this.shadowRoot.querySelectorAll('section');
        sections.forEach(section => {
          // Perform post render operations per section
          this.handleTokenOverflow(section);
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
        overflowToken.classList.add('token-list__item--show-overflow');
        overflowToken.tabIndex = 0;
        overflowToken.innerText = `+ ${overflowItems.length} more`;

        // On click or enter(accessibility), remove [+n more] token and reveal the hidden overflow tokens.
        const revealOverflow = () => {
          overflowToken.remove();
          overflowItems.forEach(i => i.classList.remove('overflow'));
        };
        overflowToken.addEventListener('click', (e: MouseEvent) => {
          revealOverflow();
        });
        overflowToken.addEventListener('keydown', (e: KeyboardEvent) => {
          if (e.code === 'Enter') {
            revealOverflow();
          }
        });
        tokenList.appendChild(overflowToken);
      }
    }
  }
}
