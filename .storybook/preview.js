/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/* global window */

import { setCustomElements } from '@storybook/web-components';
import '../node_modules/@webcomponents/webcomponentsjs/webcomponents-bundle.js';
import customElements from '../custom-elements.json';
import { versionInfo } from './versionInfo';
import { defaultDocsPage } from './story-elements/defaultDocsPage';
import './preview.css';

const setCustomElementsManifestWithOptions = (customElements, options) => {
  let { privateFields = true } = options;
  if (!privateFields) {
    customElements?.modules?.forEach(module => {
      module?.declarations?.forEach(declaration => {
        Object.keys(declaration).forEach(key => {
          if (Array.isArray(declaration[key])) {
            declaration[key] = declaration[key].filter(
              member => !member.privacy?.includes('private') && !member.privacy?.includes('protected')
            );
          }
        });
      });
    });
  }
  return setCustomElements(customElements);
};

setCustomElementsManifestWithOptions(customElements, { privateFields: false });

export const parameters = {
  previewTabs: {
    'storybook/docs/panel': {
      hidden: true
    },
    canvas: {
      hidden: true
    }
  },
  docs: {
    page: defaultDocsPage,
    story: {
      inline: false,
      height: '500px'
    }
  },
  version: versionInfo,
  options: {
    storySort: {
      order: ['stories']
    }
  }
};

const req = require.context('../stories', true, /\.(js|mdx)$/);
// configure(req, module);
if (module.hot) {
  module.hot.accept(req.id, () => {
    const currentLocationHref = window.location.href;
    window.history.pushState(null, null, currentLocationHref);
    window.location.reload();
  });
}
