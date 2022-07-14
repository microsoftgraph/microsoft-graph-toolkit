/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Graph Explorer',
  decorators: [withCodeEditor]
};

export const PersonMe = () => html`
  <mgt-person person-query="me" view="twolines" show-presence person-card="hover"></mgt-person>
`;

export const PersonUserId = () => html`
  <mgt-person user-id="{entityId}" view="twolines" show-presence person-card="hover"></mgt-person>
`;

export const People = () => html`<!-- Display the 5 most relevant people around you -->
<mgt-people show-max="5"></mgt-people>
`;

export const Groups = () => html`<!-- Display the 5 most relevant people for the specified group -->
<mgt-people group-id="{entityId}" show-max="5"></mgt-people>

<!-- Allows picking users from the specified group -->
<mgt-people-picker group-id="{entityId}"></mgt-people>
`;

export const PlannerTasks = () => html`<!-- Display your Planner tasks -->
<mgt-tasks></mgt-tasks>
`;

export const TodoTasks = () => html`<!-- Display your To Do tasks -->
<mgt-todo></mgt-todo>
`;

export const GetSingle = data => {
  function test() {
    console.log('TEST!!!!');
  }

  return html`
<style>
table {
  font-family: "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif;
  overflow-x: clip;
}

table {
  border-width: 0px;
  width: 100%;
  background-color: #FFFFFF;
  border-collapse: collapse;
  border-style: solid;
  color: #000000;
}

pre {
  white-space: pre-wrap;
}

table td, table th {
  padding-bottom: 8px;
  padding-left: 8px;
  padding-right: 20px;
  padding-top: 8px;
  text-align: left;
  font-size: 14px;
  font-weight: 400;
}

table thead > tr  {
  border-bottom-color: #323130;
}

table tr  {
  border-bottom-width: 1px;
  border-bottom-style: solid;
  border-bottom-color: rgb(200, 198, 196)
}

table th {
  padding: 5px;
  text-align: left;
  font-size: 14px;
  font-weight: 600;
  color:rgb(96, 94, 92);
}
</style>

<mgt-get resource="{resource}" max-pages="1" version="{version}">
    <template>
        <table>
            <thead>
                <tr>
                    <th>Key</th>
                    <th>Value</th>
                </tr>
            </thead>
            <tbody>
                {dynamicTemplate}
            </tbody>
        </table>
    </template>
    <template data-type="loading">
        Loading...
    </template>
    <template data-type="error">
        <pre>{{ this }}</pre>
    </template>
</mgt-get>
`;
};

GetSingle.prototype.render = () => {
  console.log('RENDER');
};

export const GetCollection = () => html`
<style>
  .template {
    font-family: "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif;
    overflow-x: clip;
  }
  table {
    border-width: 0px;
    width: 100%;
    background-color: #FFFFFF;
    border-collapse: collapse;
    border-style: solid;
    color: #000000;
  }

  pre {
    white-space: pre-wrap;
  }

  table td, table th {
    padding-bottom: 8px;
    padding-left: 8px;
    padding-right: 20px;
    padding-top: 8px;
    text-align: left;
    font-size: 14px;
    font-weight: 400;
  }

  table thead > tr  {
    border-bottom-color: #323130;
  }

  table tr  {
    border-bottom-width: 1px;
    border-bottom-style: solid;
    border-bottom-color: rgb(200, 198, 196)
  }

  table th {
    padding: 5px;
    text-align: left;
    font-size: 14px;
    font-weight: 600;
    color:rgb(96, 94, 92);
  }

  .break {
    word-break: break-all;
  }
</style>

<mgt-get resource="{resource}" max-pages="1" version="{version}">
    <template data-type="value">
        <div class="template">
            <h3 class="break">{{ this.id }}</h3>
            <table>
                <thead>
                    <tr>
                        <th>Key</th>
                        <th>Value</th>
                    </tr>
                </thead>
                <tbody>
                    <tr data-for='key in Object.keys(this)'>
                        <td>
                            <pre data-if='!key.startsWith("@")'>{{ this.{{ key }} }}</pre>
                            <pre data-else>{{ this["{{ key }}"] }}</pre>
                        </td>
                        <td class="break" data-if="this[key] === Object(this[key])">
                            <pre>{{ this[key] }}</pre>
                        </td>
                        <td class="break" data-else>{{ this[key] }}</td>
                    </tr>
                </tbody>
            </table>
        </div>
    </template>
    <template data-type="loading">
        Loading...
    </template>
    <template data-type="error">
        <pre>{{ this }}</pre>
    </template>
</mgt-get>
`;
