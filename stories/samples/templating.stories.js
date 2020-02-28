/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withSignIn } from '../../.storybook/addons/signInAddon/signInAddon';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';
import '../../dist/es6/components/mgt-get/mgt-get';

export default {
  title: 'Samples | Templating',
  component: 'mgt-get',
  decorators: [withSignIn, withCodeEditor],
  parameters: {
    signInAddon: {
      test: 'test'
    }
  }
};

export const PersonCardAdditionalDetails = () => html`
  <mgt-person person-query="me" show-name show-email person-card="hover">
    <template data-type="person-card">
      <mgt-person-card inherit-details>
        <template data-type="additional-details">
          <h3>Stuffed Animal Friends:</h3>
          <ul>
            <li>Giraffe</li>
            <li>lion</li>
            <li>Rabbit</li>
          </ul>
        </template>
      </mgt-person-card>
    </template>
  </mgt-person>
  <div style="margin:2em 0 0 1em;font-family:segoe ui;color:#323130;font-size:12px">
    (Hover on person to view Person Card)
  </div>
`;

export const GroupedEmail = () => html`
  <mgt-get resource="/me/messages" scopes="mail.read" max-pages="2">
    <template>
      <div>
        <div data-for="group in groupMail(value)">
          <div class="header">
            <mgt-person person-query="{{ group[0] }}" show-name person-card="hover"></mgt-person>
          </div>
          <div data-for="message in group[1]" class="email">
            <h2>{{ message.subject }}</h2>
            <b>Preview:</b> {{ message.bodyPreview }}
          </div>
        </div>
      </div>
    </template>
  </mgt-get>

  <script>
    document.querySelector('mgt-get').templateContext = {
      groupMail: messages => {
        let groupBy = (list, keyGetter) => {
          const map = new Map();
          list.forEach(item => {
            const key = keyGetter(item);
            const collection = map.get(key);
            if (!collection) {
              map.set(key, [item]);
            } else {
              collection.push(item);
            }
          });
          return map;
        };

        let grouped = groupBy(messages, m => m.sender.emailAddress.address);
        return [...grouped];
      }
    };
  </script>

  <style>
    .header {
      background-color: lightblue;
      padding: 20px 10px;
      margin: 8px 4px;
      box-shadow: 0 3px 7px rgba(0, 0, 0, 0.3);
    }

    .email {
      padding: 10px;
      margin: 8px 16px;
      font-family: Segoe UI, Frutiger, Frutiger Linotype, Dejavu Sans, Helvetica Neue, Arial, sans-serif;
    }

    .email:hover {
      box-shadow: 0 3px 14px rgba(0, 0, 0, 0.3);
    }

    .email h2 {
      font-size: 10px;
      margin-top: 0px;
      margin-bottom: 0px;
    }
  </style>
`;
