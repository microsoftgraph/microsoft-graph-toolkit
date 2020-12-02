/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withSignIn } from '../../.storybook/addons/signInAddon/signInAddon';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Samples | General',
  component: 'mgt-combo',
  decorators: [withSignIn, withCodeEditor],
  parameters: {
    a11y: {
      disabled: true
    },
    signInAddon: {
      test: 'test'
    }
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
          sendEmailLinkSubtitle: 'ارسل بريد الكتروني',
          startChatLinkSubtitle: 'ابدأ الدردشة',
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

export const theme = () => html`
  <div class="mgt-light">
    <header class="mgt-dark">
      <p>I should be dark, regional class</p>
      <mgt-teams-channel-picker></mgt-teams-channel-picker>
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
