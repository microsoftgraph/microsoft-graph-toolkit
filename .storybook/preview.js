/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/* global window */

import { addParameters, setCustomElements, configure } from '@storybook/web-components';
import { addons } from '@storybook/addons';
import '../packages/mgt';
import '../node_modules/@webcomponents/webcomponentsjs/webcomponents-bundle.js';
import theme from './theme';

addParameters({
  docs: {
    iframeHeight: '400px',
    inlineStories: false
  }
});

addons.setConfig({
  theme,
  enableShortcuts: false
});

// force full reload to not reregister web components
const req = require.context('../stories', true, /\.(js|mdx)$/);
configure(req, module);
if (module.hot) {
  module.hot.accept(req.id, () => {
    const currentLocationHref = window.location.href;
    window.history.pushState(null, null, currentLocationHref);
    window.location.reload();
  });
}
