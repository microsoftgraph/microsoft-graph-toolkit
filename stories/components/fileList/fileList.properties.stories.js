/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-file-list / Properties',
  component: 'file-list',
  decorators: [withCodeEditor]
};

export const getFileListByListQuery = () => html`
    <mgt-file-list
      file-list-query="/me/drive/items/01BYE5RZYJ43UXGBP23BBIFPISHHMCDTOY/children">
    </mgt-file-list>
  `;

export const getFileListByItemId = () => html`
    <mgt-file-list item-id="01BYE5RZYJ43UXGBP23BBIFPISHHMCDTOY"></mgt-file-list>
  `;

export const getFileListByItemPath = () => html`
    <mgt-file-list
      item-path="/Class%20Documents"
    ></mgt-file-list>
  `;

export const getFileListByFiles = () => html`
    <mgt-file-list
      files="01OXYKUGW6HG5WM2OBTVDZ72ABJAY5P4BY,01BYE5RZZFWGWWVNHHKVHYXE3OUJHGWCT2"
    ></mgt-file-list>
  `;

export const getFileListByFileQueries = () => html`
    <mgt-file-list
      file-queries="/me/drive/items/01BYE5RZZFWGWWVNHHKVHYXE3OUJHGWCT2,/me/drive/items/01BYE5RZ5MYLM2SMX75ZBIPQZIHT6OAYPB,/me/drive/items/01BYE5RZ47DTJGHO73WBH2ONNXQZVNNILJ"
    ></mgt-file-list>
  `;

export const getFileListByGroupId = () => html`
    <mgt-file-list
      group-id="8090c93e-ba7c-433e-9f39-08c7ba07c0b3"
      item-id="01AYQNNE76S6ES2SZFKFEKVD77I7JBARMB"
    ></mgt-file-list>
    <mgt-file-list group-id="8090c93e-ba7c-433e-9f39-08c7ba07c0b3" item-path="/Design"></mgt-file-list>
  `;

export const getFileListByDriveId = () => html`
    <mgt-file-list
      drive-id="b!-RIj2DuyvEyV1T4NlOaMHk8XkS_I8MdFlUCq1BlcjgmhRfAj3-Z8RY2VpuvV_tpd"
      item-id="01BYE5RZYJ43UXGBP23BBIFPISHHMCDTOY"
    ></mgt-file-list>
    <mgt-file-list
      drive-id="b!-RIj2DuyvEyV1T4NlOaMHk8XkS_I8MdFlUCq1BlcjgmhRfAj3-Z8RY2VpuvV_tpd"
      item-path="/Class%20Documents"
    ></mgt-file-list>
  `;

export const getSignedinUserFileList = () => html`
    <mgt-file-list item-id="01BYE5RZYJ43UXGBP23BBIFPISHHMCDTOY"></mgt-file-list>
  `;

export const getFileListBySiteId = () => html`
    <mgt-file-list
      site-id="m365x214355.sharepoint.com,5a58bb09-1fba-41c1-8125-69da264370a0,9f2ec1da-0be4-4a74-9254-973f0add78fd"
      item-id="01OXYKUGW6HG5WM2OBTVDZ72ABJAY5P4BY"
    ></mgt-file-list>
    <mgt-file-list
      site-id="m365x214355.sharepoint.com,5a58bb09-1fba-41c1-8125-69da264370a0,9f2ec1da-0be4-4a74-9254-973f0add78fd"
      item-Path="/DemoDocs/AdminDemo"
    ></mgt-file-list>
  `;

export const getFileListByUserId = () => html`
    <mgt-person user-id="48d31887-5fad-4d73-a9f5-3c356e68a038" view="twolines"></mgt-person>
    <mgt-file-list
      user-id="48d31887-5fad-4d73-a9f5-3c356e68a038"
      item-id="01BYE5RZYFPM65IDVARFELFLNTXR4ZKABD"
    ></mgt-file-list>
    <mgt-file-list user-id="48d31887-5fad-4d73-a9f5-3c356e68a038" item-path="Contoso Electronics"></mgt-file-list>
  `;

export const getFileListByInsights = () => html`
    <mgt-file-list user-id="e3d0513b-449e-4198-ba6f-bd97ae7cae85" insight-type="trending"></mgt-file-list>
    <mgt-file-list insight-type="shared"></mgt-file-list>
  `;

export const getFileListByExtensions = () => html`
    <mgt-file-list file-extensions="docx, xlsx"></mgt-file-list>
  `;

export const disableOpenOnClick = () => html`
<mgt-file-list disable-open-on-click></mgt-file-list>
`;

export const disableFileExpansion = () => html`
    <mgt-file-list hide-more-files-button></mgt-file-list>
  `;

export const getFileListWithSize = () => html`
    <mgt-file-list page-size=5></mgt-file-list>
  `;

export const getFileListByExtensionsAndSize = () => html`
    <mgt-file-list file-extensions="docx, xlsx" page-size=5></mgt-file-list>
  `;

export const fileListItemView = () => html`
    <mgt-file-list item-view="oneline" page-size=3></mgt-file-list>
    <mgt-file-list item-view="twolines" page-size=3></mgt-file-list>
    <mgt-file-list item-view="threelines" page-size=3></mgt-file-list>
  `;

export const clearCacheAndReload = () => html`
  <button>Reload files!</button>
  <mgt-file-list></mgt-file-list>
  <script>
    const fileList = document.querySelector('mgt-file-list');
    document.querySelector('button').addEventListener('click', () => {
      // passing true will clear file cache before reloading
      fileList.reload(true);
      alert("Files Reloaded");
    })
  </script>
`;

export const fileListUpload = () => html`
    <mgt-file-list enable-file-upload></mgt-file-list>
    <!-- Include a maximun file size -->
    <mgt-file-list max-file-size="10000" enable-file-upload></mgt-file-list>
  `;
