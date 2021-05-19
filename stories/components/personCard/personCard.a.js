/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components | mgt-person-card',
  component: 'mgt-person-card',
  decorators: [withCodeEditor]
};

export const personCard = () => html`
  <mgt-person-card person-query="me"></mgt-person-card>
`;

export const events = () => html`
  <!-- Open dev console and click on an event -->
  <!-- See js tab for event subscription -->

  <mgt-person-card person-query="me"></mgt-person-card>
  <script>
    const personCard = document.querySelector('mgt-person-card');
    personCard.addEventListener('expanded', () => {
      console.log("expanded");
    })
  </script>
`;

export const RTL = () => html`
  <body dir="rtl">
    <mgt-person-card person-query="me"></mgt-person-card>
  </body>
`;

export const localization = () => html`
  <mgt-person-card person-query="me"></mgt-person-card>
  <script>
  import { LocalizationHelper } from '@microsoft/mgt';
  LocalizationHelper.strings = {
    _components: {
      login: {
        signInLinkSubtitle: 'تسجيل الدخول',
        signOutLinkSubtitle: 'خروج'
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
