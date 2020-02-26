/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withSignIn } from '../.storybook/addons/signInAddon/signInAddon';
import { withCodeEditor } from '../.storybook/addons/codeEditorAddon/codeAddon';
import '../dist/es6/components/mgt-get/mgt-get';

export default {
  title: 'Samples | mgt-get',
  component: 'mgt-get',
  decorators: [withSignIn, withCodeEditor],
  parameters: {
    signInAddon: {
      test: 'test'
    }
  }
};

export const DetailedPersonCard = () => html`
  <mgt-person person-query="Isaiah" show-name show-email person-card="hover">
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

export const GetEmails = () => html`
	<mgt-get resource="/me/messages" version="beta" scopes="mail.read" max-pages="2">
		<template>
			Emails: {{value.length}}
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
			loading...
		</template>
	</mgt-get>
`;
