/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';
import { defaultDocsPage } from '../../../.storybook/story-elements/defaultDocsPage';

export default {
  title: 'Components / mgt-file-list',
  component: 'file-list',
  decorators: [withCodeEditor],
  parameters: {
    docs: {
      page: defaultDocsPage,
      source: { code: '<mgt-file-list></mgt-file-list>' }
    }
  }
};

export const fileList = () => html`
  <mgt-file-list></mgt-file-list>
`;

export const RTL = () => html`
  <body dir="rtl">
    <mgt-file-list></mgt-file-list>
  </body>
`;

export const localization = () => html`
  <mgt-file-list></mgt-file-list>
  <script>
  import { LocalizationHelper } from '@microsoft/mgt-element';
  LocalizationHelper.strings = {
    _components: {
      'file-list': {
        showMoreSubtitle: 'Show more 📂'
      },
      'file': {
        modifiedSubtitle: '⚡',
      }
    }
  }
  </script>
`;

export const events = () => html`
  <p>Clicked File:</p>
  <mgt-file></mgt-file>
  <mgt-file-list disable-open-on-click></mgt-file-list>
  <script>
    document.querySelector('mgt-file-list').addEventListener('itemClick', e => {
      const file = document.querySelector('mgt-file');
      file.fileDetails = e.detail;
    });
  </script>
  <style>
    body {
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto,
        'Helvetica Neue', sans-serif;
    }

    p {
      margin: 0;
    }
  </style>
`;

export const openFolderBreadcrumbs = () => html`
    <style>
      body {
        font-family: 'Segoe UI', 'Segoe UI Web (West European)', 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto,
          'Helvetica Neue', sans-serif;
      }

      ul.breadcrumb {
        margin: 0;
        padding: 10px 16px;
        list-style: none;
        background-color: var(--neutral-layer-1);
        font-size: 12px;
      }

      ul.breadcrumb li {
        display: inline;
      }

      ul.breadcrumb li + li:before {
        padding: 8px;
        color: var(--foreground-on-accent-fill);
        content: '\/';
      }

      ul.breadcrumb li a {
        color: var(--accent-fill-rest);
        text-decoration: none;
      }

      ul.breadcrumb li a:hover {
        color: var(--accent-fill-hover);
        text-decoration: underline;
      }
    </style>

    <ul class="breadcrumb" id="nav">
      <li><a id="home">Files</a></li>
    </ul>
    <mgt-file-list></mgt-file-list>

    <script type="module">
      const fileList = document.querySelector('mgt-file-list');
      const nav = document.getElementById('nav');
      const home = document.getElementById('home');

      let homeListId;
      if (fileList.itemId) {
        homeListId = fileList.itemId;
      } else {
        homeListId = null;
      }

      // handle default file list menu item
      home.addEventListener('click', e => {
        fileList.itemId = homeListId;
        removeListItems(1);
      });

      // handle create and remove menu items
      fileList.addEventListener('itemClick', e => {
        if (e.detail && e.detail.folder) {
          const id = e.detail.id;
          const name = e.detail.name;

          // render new file list
          fileList.itemId = id;

          // create breadcrumb menu item
          const li = document.createElement('li');
          const a = document.createElement('a');
          li.setAttribute('id', id);
          a.appendChild(document.createTextNode(name));
          li.appendChild(a);
          nav.appendChild(li);

          // remove breadcrumb menu items and render file list based on clicked item
          a.addEventListener('click', e => {
            const nodes = Array.from(nav.children);
            const index = nodes.indexOf(li);
            if (e.target) {
              removeListItems(index + 1);
              fileList.itemId = li.id;
            }
          });
        }
      });

      // remove li of ul where index is larger than n
      function removeListItems(n) {
        while (nav.getElementsByTagName('li').length > n) {
          nav.removeChild(nav.lastChild);
        }
      }
    </script>
  `;
