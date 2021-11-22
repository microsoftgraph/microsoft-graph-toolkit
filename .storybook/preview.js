/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/* global window */

import { addParameters } from '@storybook/web-components';
import '../packages/mgt';
import '../node_modules/@webcomponents/webcomponentsjs/webcomponents-bundle.js';
import theme from './theme';

addParameters({
  docs: {
    iframeHeight: '400px',
    inlineStories: false
  }
});
