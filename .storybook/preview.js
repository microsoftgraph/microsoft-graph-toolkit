/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/* global window */

import { configure, addParameters, setCustomElements } from '@storybook/web-components';
import customElements from '../custom-elements.json';
import '../node_modules/@webcomponents/webcomponentsjs/webcomponents-bundle.js';
import theme from './theme';

setCustomElements(customElements);

addParameters({
  docs: {
    iframeHeight: '400px',
    inlineStories: false
  },
  options: {
    // disable keyboard shortcuts because they interfere with the stories
    enableShortcuts: false,
    theme
  }
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
