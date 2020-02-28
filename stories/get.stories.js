/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withA11y } from '@storybook/addon-a11y';
import { withKnobs } from '@storybook/addon-knobs';
import { withWebComponentsKnobs } from 'storybook-addon-web-components-knobs';
import { withSignIn } from '../.storybook/addons/signInAddon/signInAddon';
import { withCodeEditor } from '../.storybook/addons/codeEditorAddon/codeAddon';
import '../dist/es6/components/mgt-get/mgt-get';

export default {
  title: 'mgt-get',
  component: 'mgt-get',
  decorators: [withA11y, withSignIn, withCodeEditor],
  parameters: {
    options: { selectedPanel: 'mgt/sign-in' },
    signInAddon: {
      test: 'test'
    }
  }
};

export const Email = () => html`
  <mgt-get resource="/me/messages" version="beta" scopes="mail.read" max-pages="2">
      <template>
        emails: {{value.length}}
        <ol>
          <li data-for="email in value">
            <div>
              <h2>{{ email.subject }}</h2>
              <span>
                <b>From:</b> <mgt-person
                person-query="{{ email.sender.emailAddress.address }}"
                show-name
                person-card="hover"
                ></mgt-person>
              </span>
              <br />
              <b>Preview:</b> {{ email.bodyPreview }}
            </div>
          </li>
        </ul>
      </template>

      <template data-type="error">
        {{ this }}
      </template>

      <template data-type="loading">
        loading
      </template>
    </mgt-get>

    <script>
      let get = document.querySelector('mgt-get');
      get.addEventListener('dataChange', e => {
        console.log('data changed');
        console.log('response ', get.response);
        console.log('error ', get.error);
      })
    </script>
`;

export const GroupedEmail = () => html`
  <mgt-get resource="/me/messages" scopes="mail.read" max-pages="2">
    <template>
      <div>
        <div data-for="group in groupMail(value)">
          <mgt-person person-query="{{ group[0] }}" show-name person-card="hover"></mgt-person>
          <div data-for="message in group[1]">
            <h2>{{ message.subject }}</h2>
            <b>Preview:</b> {{ message.bodyPreview }}
          </div>
        </div>
      </div>
    </template>
  </mgt-get>

  <script>
    const get = document.querySelector('mgt-get');
    get.templateContext = get.templateContext || {};
    get.templateContext.groupMail = messages => {
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
    };
  </script>
`;

export const TeamsMessaging = () => html`
  <mgt-get resource="/me/joinedTeams" id="teamsGet">
    <template>
      Team:
      <select data-props="{{@change: teamChange}}">
        <option data-for="{{team in value}}" value="{{team.id}}">{{team.displayName}}</option>
      </select>
    </template>
  </mgt-get>

  <mgt-get id="channelsGet">
    <template>
      Channel:
      <select data-props="{{@change: channelChange}}">
        <option data-for="channel in value" value="{{channel.id}}">{{channel.displayName}}</option>
      </select>
    </template>
    <template data-type="loading">
      loading
    </template>
  </mgt-get>

  <mgt-get id="messagesGet" version="beta" polling-rate="3000">
    <template data-type="value">
      <div data-if="!deletedDateTime" class="teams-message">
        <mgt-person user-id="{{from.user.id}}" show-name></mgt-person>
        <div>{{body.content}}</div>
      </div>
    </template>
    <template data-type="loading">
      loading
    </template>
  </mgt-get>

  <div id="sendMessageDiv" class="hidden">
    <input id="messageText" type="text" />
    <button id="sendMessageButton">Send message</button>
  </div>

  <script type="module">
    import { Providers } from '../../../dist/es6/index.js';

    const teamsGet = document.getElementById('teamsGet');
    const channelsGet = document.getElementById('channelsGet');
    const messagesGet = document.getElementById('messagesGet');

    let teamId, channelId;

    const setTeamId = id => {
      teamId = id;
      channelsGet.resource = teamId ? \`teams/\${teamId}/channels\` : null;
      messagesGet.resource = null;
    };

    const setChannelId = id => {
      channelId = id;
      messagesGet.resource = channelId ? \`teams/\${teamId}/channels/\${channelId}/messages/delta\` : null;
    };

    teamsGet.templateContext = teamsGet.templateContext || {};
    teamsGet.templateContext['teamChange'] = e => {
      setTeamId(e.target.options[e.target.selectedIndex].value);
    };

    channelsGet.templateContext = channelsGet.templateContext || {};
    channelsGet.templateContext['channelChange'] = e => {
      setChannelId(e.target.options[e.target.selectedIndex].value);
    };

    teamsGet.addEventListener('dataChange', e => {
      setTeamId(e.detail.response ? e.detail.response.value[0].id : null);
    });

    channelsGet.addEventListener('dataChange', e => {
      setChannelId(e.detail.response ? e.detail.response.value[0].id : null);
    });

    messagesGet.addEventListener('dataChange', e => {
      if (e.detail.response) {
        document.querySelector('#sendMessageDiv').className = '';
      } else {
        document.querySelector('#sendMessageDiv').className = 'hidden';
      }

      setTimeout(() => {
        messagesGet.scrollTop = messagesGet.scrollHeight;
      }, 100);
    });

    document.querySelector('#sendMessageButton').addEventListener('click', async e => {
      const messageInput = document.querySelector('#messageText');
      const messageToSend = messageInput.value;

      const content = {
        body: {
          contentType: 'text',
          content: messageToSend
        }
      };

      await Providers.globalProvider.graph.client
        .api(\`/teams/\${teamId}/channels/\${channelId}/messages\`)
        .version('beta')
        .post(content);

      messageInput.value = '';
    });
  </script>

  <style>
    .teams-message {
      box-shadow: 0 3px 7px rgba(0, 0, 0, 0.3);
      padding: 10px;
      margin: 8px 4px;
      font-family: Segoe UI, Frutiger, Frutiger Linotype, Dejavu Sans, Helvetica Neue, Arial, sans-serif;
    }

    .teams-message:hover {
      box-shadow: 0 3px 14px rgba(0, 0, 0, 0.3);
    }

    .hidden {
      display: none;
    }

    #messagesGet {
      max-height: 300px;
      overflow: auto;
      display: block;
    }
  </style>
`;

export const Polling = () => html`
  <mgt-get resource="/me/presence" version="beta" scopes="Presence.Read" polling-rate="5000">
    <template>
      {{availability}}
    </template>
    <template data-type="error">
      {{this}}
    </template>
  </mgt-get>
`;
