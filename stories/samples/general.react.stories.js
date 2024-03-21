/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Samples / General / React',
  component: 'person-card',
  decorators: [withCodeEditor]
};

export const Hooks = () => html`
  <p>This demonstrates how to use the <code>useIsSignedIn</code> hook in React.</p>
  <mgt-person-card person-query="me"></mgt-person-card>
  <react>
    import { PersonCard, useIsSignedIn } from '@microsoft/mgt-react';
    const [isSignedIn] = useIsSignedIn();

    export default () => (
      isSignedIn ? <PersonCard personQuery="me"></PersonCard> : null
    );
  </react>
`;
