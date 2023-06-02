/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Samples / General',
  decorators: [withCodeEditor],
  parameters: {
    viewMode: 'story'
  }
};

export const LoginToShowAgenda = () => html`
  <mgt-login></mgt-login>
  <mgt-agenda></mgt-agenda>
`;

export const Localization = () => html`
  <mgt-login></mgt-login>
  <mgt-people-picker></mgt-people-picker>
  <mgt-teams-channel-picker></mgt-teams-channel-picker>
  <mgt-tasks></mgt-tasks>
  <mgt-agenda></mgt-agenda>
  <mgt-people></mgt-people>
  <mgt-todo></mgt-todo>
  <script>
    import { LocalizationHelper } from '@microsoft/mgt';
    LocalizationHelper.strings = {
      _components: {
        login: {
          signInLinkSubtitle: 'تسجيل الدخول',
          signOutLinkSubtitle: 'خروج'
        },
        'people-picker': {
          inputPlaceholderText: 'ابدأ في كتابة الاسم',
          noResultsFound: 'لم نجد أي قنوات',
          loadingMessage: '...جار التحميل'
        },
        'teams-channel-picker': {
          inputPlaceholderText: 'حدد قناة',
          noResultsFound: 'لم يتم العثور على نتائج',
          loadingMessage: 'Loading...'
        },
        tasks: {
          removeTaskSubtitle: 'delete',
          cancelNewTaskSubtitle: 'canceltest',
          newTaskPlaceholder: 'newTaskTest',
          addTaskButtonSubtitle: 'addme'
        },
        todo: {
          removeTaskSubtitle: 'todoremoveTEST'
        },
        'person-card': {
          showMoreSectionButton: 'أظهر المزيد' // global declaration
        },
        'person-card-contact': {
          contactSectionTitle: 'اتصل'
        },
        'person-card-organization': {
          reportsToSectionTitle: 'تقارير ل',
          directReportsSectionTitle: 'تقارير مباشرة',
          organizationSectionTitle: 'منظمة',
          youWorkWithSubSectionTitle: 'انت تعمل مع',
          userWorksWithSubSectionTitle: 'يعمل مع'
        },
        'person-card-messages': {
          emailsSectionTitle: 'رسائل البريد الإلكتروني'
        },
        'person-card-files': {
          filesSectionTitle: 'الملفات',
          sharedTextSubtitle: 'مشترك'
        },
        'person-card-profile': {
          SkillsAndExperienceSectionTitle: 'المهارات والخبرة',
          AboutCompactSectionTitle: 'حول',
          SkillsSubSectionTitle: 'مهارات',
          LanguagesSubSectionTitle: 'اللغات',
          WorkExperienceSubSectionTitle: 'خبرة في العمل',
          EducationSubSectionTitle: 'التعليم',
          professionalInterestsSubSectionTitle: 'المصالح المهنية',
          personalInterestsSubSectionTitle: 'اهتمامات شخصية',
          birthdaySubSectionTitle: 'عيد الميلاد',
          currentYearSubtitle: 'السنة الحالية'
        }
      }
    };
  </script>
`;

export const cache = () => html`
  <button id="ClearCacheButton" type="button">Clear Cache</button>
  <span class="notes"
    >*Note* Please refer to your browser Developer Tools -> Applications -> Storage -> IndexedDB for cached
    objects</span
  >

  <mgt-login></mgt-login>
  <mgt-person person-query="me" view="twoLines" person-card="hover" show-presence></mgt-person>
  <mgt-person user-id="4782e723-f4f4-4af3-a76e-25e3bab0d896" view="twoLines"></mgt-person>
  <mgt-people-picker></mgt-people-picker>
  <mgt-people-picker type="group"></mgt-people-picker>
  <mgt-people-picker group-id="02bd9fd6-8f93-4758-87c3-1fb73740a315"></mgt-people-picker>
  <mgt-teams-channel-picker></mgt-teams-channel-picker>
  <mgt-tasks data-source="todo"></mgt-tasks>
  <mgt-agenda group-by-day></mgt-agenda>
  <mgt-people show-presence show-max="10"></mgt-people>
  <mgt-people group-id="02bd9fd6-8f93-4758-87c3-1fb73740a315"></mgt-people>
  <mgt-people user-ids="4782e723-f4f4-4af3-a76e-25e3bab0d896, f5289423-7233-4d60-831a-fe107a8551cc"></mgt-people>
  <style>
    .notes {
      display: block;
      font-size: 12px;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto,
        'Helvetica Neue', sans-serif;
      margin-bottom: 5px;
    }
  </style>
  <script>
    import { CacheService } from '@microsoft/mgt';
    CacheService.config.isEnabled = true;

    let clearCacheButton = document.getElementById('ClearCacheButton');
    clearCacheButton.addEventListener('click', clearCache);
    function clearCache() {
      CacheService.clearCaches();
    }

    // you can clear cache by:
    // CacheService.clearCaches();

    // this is the config object
    // {
    //   defaultInvalidationPeriod: number,
    //   isEnabled: boolean,
    //   people: {
    //     invalidationPeriod: number,
    //     isEnabled: boolean
    //   },
    //   photos: {
    //     invalidationPeriod: number,
    //     isEnabled: boolean
    //   },
    //   users: {
    //     invalidationPeriod: number,
    //     isEnabled: boolean
    //   },
    //   presence: {
    //     invalidationPeriod: number,
    //     isEnabled: boolean
    //   },
    //   groups: {
    //     invalidationPeriod: number,
    //     isEnabled: boolean
    //   }
    // };
  </script>
`;
export const theme = () => html`
  <div class="mgt-light root">
    <header class="mgt-dark">
      <p>I should be dark, regional class</p>
      <mgt-people-picker></mgt-people-picker>
      <div class="mgt-light">
        <p>I should be light, second level regional class</p>
        <mgt-teams-channel-picker></mgt-teams-channel-picker>
      </div>
    </header>
    <article>
      <p>I should be light, global class</p>
      <mgt-teams-channel-picker></mgt-teams-channel-picker>
    </article>
    <p>I am custom themed</p>
    <mgt-teams-channel-picker class="custom1"></mgt-teams-channel-picker>
    <p>I have both custom input background color and mgt-dark theme</p>
    <mgt-teams-channel-picker class="mgt-dark custom2"></mgt-teams-channel-picker>
    <p>I should be light, with unknown class mgt-foo</p>
    <mgt-teams-channel-picker class="mgt-foo"></mgt-teams-channel-picker>
  </div>
  <style>
    .root {
      --focus-ring-color: red;
      --focus-ring-style: solid;
    }
    .custom1 {
      --input-border: 2px solid teal;
      --input-background-color: #33c2c2;
      --dropdown-background-color: #33c2c2;
      --dropdown-item-hover-background: #2a7d88;
      --input-hover-color: #b911b1;
      --input-focus-color: #441540;
      --font-color: white;
      --placeholder-default-color: white;
      --placeholder-focus-color: #441540;
    }

    .custom2 {
      --input-background-color: #e47c4d;
    }
  </style>
`;
