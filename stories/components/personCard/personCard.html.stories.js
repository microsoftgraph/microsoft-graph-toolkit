/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-person-card / HTML',
  component: 'person-card',
  decorators: [withCodeEditor]
};

export const personCard = () => html`
  <mgt-person-card person-query="me" id="online" show-presence></mgt-person-card>

  <!-- Person Card without Presence -->
  <!-- <mgt-person-card person-query="me"></mgt-person-card> -->
  <script>
    const online = {
      activity: 'Available',
      availability: 'Available',
      id: null
    };
    const onlinePerson = document.getElementById('online');
    onlinePerson.personPresence = online;
  </script>
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
    personCard.addEventListener('updated', (e) => {
      console.log("updated", e);
    });
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
  import { LocalizationHelper } from '@microsoft/mgt-element';
  LocalizationHelper.strings = {
    _components: {
      'person-card': {
        showMoreSectionButton: 'أظهر المزيد' // global declaration
      },
      'contact': {
        contactSectionTitle: 'اتصل',
        emailTitle: 'البريد الإلكتروني',
        chatTitle: 'دردشة',
        businessPhoneTitle: 'هاتف العمل',
        cellPhoneTitle: 'هاتف محمول',
        departmentTitle: ' قسم، أقسام',
        titleTitle: 'لقب',
        officeLocationTitle: 'موقع المكتب'
      },
      'organization': {
        reportsToSectionTitle: 'تقارير ل',
        directReportsSectionTitle: 'تقارير مباشرة',
        organizationSectionTitle: 'منظمة',
        youWorkWithSubSectionTitle: 'انت تعمل مع',
        userWorksWithSubSectionTitle: 'يعمل مع'
      },
      'messages': {
        emailsSectionTitle: 'رسائل البريد الإلكتروني'
      },
      'file-list': {
        filesSectionTitle: 'الملفات',
        sharedTextSubtitle: 'مشترك',
        showMoreSubtitle: 'Show more 📂'
      },
      'profile': {
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

export const AnonymousDisplay = () => html`
<div style="margin-bottom: 10px">
  <strong>Note:</strong> this story forces an anonymous context and explicity sets the user being displayed.<br />
  Refer to the JavaScript tab for setup details.
</div>
<mgt-person-card class="anonymous-display"></mgt-person-card>
<script>
  import { Providers, Msal2Provider } from './mgt.storybook.js';
  Providers.globalProvider = new Msal2Provider({ clientId: "fake" });
  const personCard = document.querySelector('.anonymous-display');
  personCard.personDetails = {
      displayName: 'Megan Bowen',
      jobTitle: 'CEO',
      userPrincipalName: 'megan@contoso.com',
      mail: 'megan@contoso.com',
      businessPhones: ['423-555-0120'],
      mobilePhone: '424-555-0130',
  };
</script>
`;
