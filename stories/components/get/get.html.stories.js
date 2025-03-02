/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-get / HTML',
  component: 'get',
  decorators: [withCodeEditor]
};

export const Get = () => html`
<mgt-get resource="/me/messages" scopes="mail.read">
  <template>
    <pre>{{ JSON.stringify(value, null, 2) }}</pre>
  </template>
</mgt-get>
`;

export const GetEmail = () => html`
  <mgt-get resource="/me/messages" version="beta" scopes="mail.read" max-pages="2">
    <template>
      <div class="email" data-for="email in value">
        <h3>{{ email.subject }}</h3>
        <mgt-person
          person-query="{{email.sender.emailAddress.address}}"
          view="oneline"
          person-card="none"
        ></mgt-person>
        <div data-if="email.bodyPreview" class="preview" innerHtml>{{email.bodyPreview}}</div>
        <div data-else class="preview">
          email body is empty
        </div>
      </div>
    </template>
    <template data-type="loading">
      loading
    </template>
    <template data-type="error">
      {{ this }}
    </template>
  </mgt-get>

  <style>
    .email {
      box-shadow: 0 3px 7px rgba(0, 0, 0, 0.3);
      padding: 10px;
      margin: 8px 4px;
      font-family: Segoe UI, Frutiger, Frutiger Linotype, Dejavu Sans, Helvetica Neue, Arial, sans-serif;
    }

    .email:hover {
      box-shadow: 0 3px 14px rgba(0, 0, 0, 0.3);
      padding: 10px;
      margin: 8px 4px;
    }

    .email h3 {
      font-size: 12px;
      margin-bottom: 4px;
    }

    .email h4 {
      font-size: 10px;
      margin-top: 0px;
      margin-bottom: 4px;
    }

    .email mgt-person {
      --font-size: 10px;
      --person-avatar-size: 12px;
    }

    .email .preview {
      font-size: 13px;
      text-overflow: ellipsis;
      word-wrap: break-word;
      overflow: hidden;
      line-height: 1.4em;
    }
  </style>
`;

export const ExtendingPersonCard = () => html`
  <mgt-person person-query="Isaiah" view="twolines" person-card="hover">
    <template data-type="person-card">
      <mgt-person-card inherit-details>
        <template data-type="additional-details">
          <mgt-get resource="/users/{{ person.id }}/profile" version="beta">
            <template>
              <div>
                <div data-if="positions && positions.length">
                  <h4>Work history</h4>
                  <div data-for="position in positions">
                    <b>{{ position.detail.jobTitle }}</b> ({{ position.detail.company.department }})
                  </div>
                  <hr />
                </div>
                <div data-if="projects && projects.length">
                  <h4>Project history</h4>
                  <div data-for="project in projects">
                    <b>{{ project.displayName }}</b>
                  </div>
                  <hr />
                </div>
                <div data-if="educationalActivities && educationalActivities.length">
                  <h4>Educational Activities</h4>
                  <div data-for="edu in educationalActivities">
                    <div data-if="edu.program.displayName"><b>program:</b> {{ edu.program.displayName }}</div>
                    <div data-if="edu.institution.displayName">
                      <b>Institution:</b> {{ edu.institution.displayName }}
                    </div>
                  </div>
                  <hr />
                </div>
                <div>
                  <h4>Interests</h4>
                  <span data-for="interest in interests">
                    {{ interest.displayName }}<span data-if="$index < interests.length - 1">, </span>
                  </span>
                  <hr />
                </div>
                <div data-if="languages && languages.length">
                  <h4>Languages</h4>
                  <span data-for="language in languages">
                    {{ language.displayName }}<span data-if="$index < languages.length - 1">, </span>
                  </span>
                </div>
              </div>
            </template>
          </mgt-get>
        </template>
      </mgt-person-card>
    </template>
  </mgt-person>
`;

export const UsingImageType = () => html`
  <mgt-get resource="me">
    <template>
      <mgt-get resource="users/{{id}}/photo/$value" type="image">
        <template>
          <mgt-person-card person-details="{{$parent}}" person-image="{{image}}"></mgt-person-card>
        </template>
      </mgt-get>
    </template>
  </mgt-get>
`;

export const UsingCaching = () => html`
  <mgt-get resource="me" cache-enabled="true">
    <template>
      Hello {{ displayName }}
    </template>
  </mgt-get>
`;

export const PollingRate = () => html`
  <mgt-get resource="/me/presence" version="beta" scopes="Presence.Read" polling-rate="2000">
    <template>
      {{availability}}
    </template>
  </mgt-get>
`;

export const refresh = () => html`
<div>
    <label>get.refresh(false)</label>
    <button id="false">Soft refresh</button>
</div>
<div>
    <label>get.refresh(true)</label>
    <button id="true">Hard refresh</button>
</div>

<mgt-get cache-enabled="true" resource="/me/messages" version="beta">
  <template data-type="default"> {{ this }}</template>
  <template data-type="loading">
    <h2>Loading...?!?!</h2>
  </template>
</mgt-get>

<script>
const softRefreshButton = document.querySelector('#false');
const hardRefreshButton = document.querySelector('#true');
const getElement = document.querySelector('mgt-get');

const softRefresh = () => {
  alert('requesting soft refresh of mgt-get component');
  getElement.refresh(false);
};

const hardRefresh = () => {
  alert('requesting hard refresh of mgt-get component');
  getElement.refresh(true);
};

softRefreshButton.addEventListener('click', softRefresh)

hardRefreshButton.addEventListener('click', hardRefresh)
</script>
`;

export const events = () => html`
  <mgt-get resource="/me/messages" scopes="mail.read">
    <template>
      <pre>{{ JSON.stringify(value, null, 2) }}</pre>
    </template>
  </mgt-get>

  <script>
    const get = document.querySelector('mgt-get');
    get.addEventListener('updated', e => {
      console.log('updated', e);
    });
    get.addEventListener('dataChange', e => {
      console.log('dataChange', e);
    });
    get.addEventListener('templateRendered', e => {
      console.log('templateRendered', e);
    });
  </script>
`;
