import { html } from 'lit-element';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';
import '../../packages/mgt-components/dist/es6/components/mgt-file-list/mgt-file-list';
import '../../packages/mgt-components/dist/es6/components/mgt-person/mgt-person';

export default {
  title: 'Components | mgt-file-list',
  component: 'mgt-file-list',
  decorators: [withCodeEditor]
};

export const fileList = () => html`
  <mgt-file-list></mgt-file-list>
`;

export const getFileListByListQuery = () => html`
  <mgt-file-list file-list-query="me/drive/items/01BYE5RZYJ43UXGBP23BBIFPISHHMCDTOY/children"></mgt-file-list>
`;

export const getFileListById = () => html`
  <mgt-file-list item-id="01BYE5RZYJ43UXGBP23BBIFPISHHMCDTOY"></mgt-file-list>
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

export const disableFileExpansion = () => html`
  <mgt-file-list disable-expansion></mgt-file-list>
`;

export const selectedFile = () => html`
  <p>Selected File:</p>
  <mgt-file></mgt-file>
  <mgt-file-list></mgt-file-list>
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

export const customCssProperties = () => html`
  <style>
    mgt-file-list {
      --file-list-background-color: #e0f8db;
      --file-item-background-color--hover: #caf1de;
      --file-item-background-color--active: #acddde;
      --file-list-border: 4px dotted #ffbdc3;
      --file-list-box-shadow: none;
      --file-list-padding: 0;
      --file-list-margin: 0;
      --file-item-border-radius: 12px;
      --file-item-margin: 2px 6px;
      --file-item-border-bottom: 4px dotted #ffbdc3;
      --show-more-button-background-color: #fef8dd;
      --show-more-button-background-color--hover: #ffe7c7;
      --show-more-button-font-size: 14px;
      --show-more-button-padding: 16px;
      --show-more-button-border-bottom-right-radius: 12px;
      --show-more-button-border-bottom-left-radius: 12px;
    }
  </style>
  <mgt-file-list></mgt-file-list>
`;

export const darkTheme = () => html`
  <mgt-file-list class="mgt-dark"></mgt-file-list>
`;
