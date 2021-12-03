/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/* global window */

import { addParameters, setCustomElements } from '@storybook/web-components';

import '../node_modules/@webcomponents/webcomponentsjs/webcomponents-bundle.js';
import customElements from '../custom-elements.json';

setCustomElements(customElements);

addParameters({
  docs: {
    iframeHeight: '400px',
    inlineStories: false
  }
});

const req = require.context('../stories', true, /\.(js|mdx)$/);
// configure(req, module);
if (module.hot) {
  module.hot.accept(req.id, () => {
    const currentLocationHref = window.location.href;
    window.history.pushState(null, null, currentLocationHref);
    window.location.reload();
  });
}
