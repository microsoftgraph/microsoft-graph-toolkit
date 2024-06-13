/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Samples / General / HTML',
  decorators: [withCodeEditor],
  parameters: {
    viewMode: 'story'
  }
};

export const Localization = () => html`
  <mgt-login></mgt-login>
  <mgt-people-picker></mgt-people-picker>
  <mgt-teams-channel-picker></mgt-teams-channel-picker>
  <mgt-planner></mgt-planner>
  <mgt-agenda></mgt-agenda>
  <mgt-people></mgt-people>
  <mgt-todo></mgt-todo>
  <script>
    import { LocalizationHelper } from '@microsoft/mgt-element';
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
        planner: {
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
        'contact': {
          contactSectionTitle: 'اتصل'
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

export const Cache = () => html`
<fluent-button id="ClearCacheButton" appearance="accent">Clear Cache</fluent-button>
<div id="status" class="notes"></div>
<span class="notes"
  >*Note* Please refer to your browser Developer Tools -> Applications -> Storage -> IndexedDB for cached
  objects</span
>

<mgt-login></mgt-login>
<mgt-person person-query="me" view="twolines" person-card="hover" show-presence></mgt-person>
<mgt-person user-id="4782e723-f4f4-4af3-a76e-25e3bab0d896" view="twolines"></mgt-person>
<mgt-people-picker></mgt-people-picker>
<mgt-people-picker type="group"></mgt-people-picker>
<mgt-people-picker group-id="02bd9fd6-8f93-4758-87c3-1fb73740a315"></mgt-people-picker>
<mgt-teams-channel-picker></mgt-teams-channel-picker>
<mgt-planner data-source="todo"></mgt-planner>
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
import { CacheService, Providers } from '@microsoft/mgt-element';
CacheService.config.isEnabled = true;

const status = document.getElementById('status');
status.innerHTML = 'Cache is enabled: ' + CacheService.config.isEnabled;

const clearCache = async () => {
  // get the id of the current cache
  const id = await Providers.getCacheId();
  status.innerHTML = 'Clearing cache ' + id + ', this may take a little while...';
  // clear the current cache
  await CacheService.clearCacheById(id);
  status.innerHTML = 'Cache cleared.';
}

const onClearCacheButtonClick = () => {
  void clearCache();
};

let clearCacheButton = document.getElementById('ClearCacheButton');
clearCacheButton.addEventListener('click', onClearCacheButtonClick);

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

export const Theme = () => html`
<div>
  <p>This demonstrates how to set the theme globally without using a theme toggle and customize styling within specific scopes</p>
  <p>Please refer to the JS and CSS tabs in the editor for implentation details</p>
  <mgt-login></mgt-login>
  <p>This picker shows the custom focus ring color</p>
  <div class="custom-focus">
    <mgt-people-picker></mgt-people-picker>
  </div>
  <article>
    <p>I use the theme set on the body</p>
    <mgt-teams-channel-picker></mgt-teams-channel-picker>
  </article>
  <p>I am custom themed, take care to ensure that your customizations maintain accessibility standards</p>
  <mgt-teams-channel-picker class="custom1"></mgt-teams-channel-picker>
</div>
<script>
import { applyTheme } from '@microsoft/mgt-components';
const body = document.querySelector('body');
if(body) applyTheme('dark', body);
</script>
<style>
body {
  background-color: var(--fill-color);
  color: var(--neutral-foreground-rest);
  font-family: var(--default-font-family, var(--body-font));
  /* Uncomment to change the default font family */
  /* --default-font-family: 'Franklin Gothic Medium', 'Arial Narrow', Arial, sans-serif; */
  /* Uncomment to change the fluent element font variable */
  /* --body-font: 'Franklin Gothic Medium', 'Arial Narrow', Arial, sans-serif; */
}
.custom-focus {
  --focus-ring-color: red;
  --focus-ring-style: solid;
}
.custom1 {
  --channel-picker-input-border: 2px solid teal;
  --channel-picker-input-background-color: black;
  --channel-picker-input-background-color-hover: #1a1a1a;
  --channel-picker-search-icon-color: yellow;
  --channel-picker-close-icon-color: yellow;
  --channel-picker-down-chevron-color: yellow;
  --channel-picker-up-chevron-color: yellow;
  --channel-picker-input-placeholder-text-color: white;
  --channel-picker-input-placeholder-text-color-focus: yellow;
  --channel-picker-input-placeholder-text-color-hover: chartreuse;
  --channel-picker-dropdown-background-color: #008383;
  --channel-picker-dropdown-item-background-color-hover: #006363;
  --channel-picker-font-color: white;
  --channel-picker-placeholder-default-color: white;
  --channel-picker-placeholder-focus-color: #441540;
}
</style>
`;

export const IsSignedIn = () => html`
  <p>This demonstrates how to use the <code>isSignedIn</code> utility function in JavaScript. If you're signed in, <code>mgt-person</code> is rendered below.</p>
  <div id="person"></div>
  <script>
    import { isSignedIn } from '@microsoft/mgt-element';
    const person = document.getElementById("person");
    if(isSignedIn() && person){
      const mgtPerson = document.createElement("mgt-person");
      mgtPerson.setAttribute("person-query", "me");
      person.appendChild(mgtPerson);
    }
</script>
`;

export const Calendar = () => html`
  <mgt-login></mgt-login></mgt-login>
  <mgt-agenda group-by-day></mgt-agenda>
`;
