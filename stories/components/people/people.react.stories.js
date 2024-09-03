/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-people / React',
  component: 'people',
  decorators: [withCodeEditor]
};

export const People = () => html`
  <mgt-people show-max="5"></mgt-people>
  <react>
    import { People } from '@microsoft/mgt-react';

    export default () => (
      <People></People>
    );
  </react>
`;

export const PeopleQueries = () => html`
  <mgt-people></mgt-people>
  <react>
    import { People } from '@microsoft/mgt-react';
    const peopleDisplay: string[] = ['LidiaH', 'Megan Bowen', 'Lynne Robbins', 'JoniS'];

    export default () => (
      <People peopleQueries={peopleDisplay}></People>
    );
  </react>
`;
