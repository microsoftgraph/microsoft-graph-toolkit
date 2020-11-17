/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withSignIn } from '../../.storybook/addons/signInAddon/signInAddon';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';
import { localize } from '../../.storybook/addons/localizeAddon/localizationAddon';
import { withKnobs } from '@storybook/addon-knobs';

export default {
  title: 'Samples | General',
  component: 'mgt-combo',
  decorators: [withSignIn, withCodeEditor, withKnobs, localize],
  parameters: {
    a11y: {
      disabled: true
    },
    signInAddon: {
      test: 'test'
    },
    strings: {
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
          noResultsFound: `لم يتم العثور على نتائج`,
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
        }
      }
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
`;
