import { html } from 'lit-element';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';
import '../../packages/mgt-components/dist/es6/components/mgt-file/mgt-file';
import '../../packages/mgt-components/dist/es6/components/mgt-person/mgt-person';

export default {
  title: 'Components | mgt-file',
  component: 'mgt-file',
  decorators: [withCodeEditor]
};

export const file = () => html`
  <mgt-file file-query="/me/drive/items/01BYE5RZZFWGWWVNHHKVHYXE3OUJHGWCT2"></mgt-file>
`;

export const fileViews = () => html`
  <mgt-file file-query="/me/drive/items/01BYE5RZZFWGWWVNHHKVHYXE3OUJHGWCT2" view="oneline"></mgt-file>
  <mgt-file file-query="/me/drive/items/01BYE5RZZFWGWWVNHHKVHYXE3OUJHGWCT2" view="twolines"></mgt-file>
  <mgt-file file-query="/me/drive/items/01BYE5RZZFWGWWVNHHKVHYXE3OUJHGWCT2" view="threelines"></mgt-file>
`;

export const setIcon = () => html`
  <mgt-file
    file-query="/me/drive/items/01BYE5RZZFWGWWVNHHKVHYXE3OUJHGWCT2"
    file-icon="https://github.com//microsoftgraph/microsoft-graph-toolkit/blob/main/assets/favicon.png?raw=true"
  ></mgt-file>
`;

export const getFileByDriveId = () => html`
  <mgt-file
    drive-id="b!-RIj2DuyvEyV1T4NlOaMHk8XkS_I8MdFlUCq1BlcjgmhRfAj3-Z8RY2VpuvV_tpd"
    item-id="01BYE5RZ5MYLM2SMX75ZBIPQZIHT6OAYPB"
  ></mgt-file>
  <mgt-file
    drive-id="b!-RIj2DuyvEyV1T4NlOaMHk8XkS_I8MdFlUCq1BlcjgmhRfAj3-Z8RY2VpuvV_tpd"
    item-path="Attachments"
  ></mgt-file>
`;

export const getFileByGroupId = () => html`
  <mgt-file group-id="02bd9fd6-8f93-4758-87c3-1fb73740a315" item-id="01XKNBVLNL4EWP43GPU5EY67UHT3DGKCWQ"></mgt-file>
  <mgt-file group-id="02bd9fd6-8f93-4758-87c3-1fb73740a315" item-path="Onboarding/Employee Handbook.docx"></mgt-file>
`;

export const getSignedinUserFiles = () => html`
  <mgt-file item-id="01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36K"></mgt-file>
  <mgt-file item-path="Class Documents/01. Organic Chemistry Header Image.jpg"></mgt-file>
`;

export const getFileBySiteId = () => html`
  <mgt-file
    site-id="m365x214355.sharepoint.com,5a58bb09-1fba-41c1-8125-69da264370a0,9f2ec1da-0be4-4a74-9254-973f0add78fd"
    item-id="01OXYKUGW6HG5WM2OBTVDZ72ABJAY5P4BY"
  ></mgt-file>
  <mgt-file
    site-id="m365x214355.sharepoint.com,5a58bb09-1fba-41c1-8125-69da264370a0,9f2ec1da-0be4-4a74-9254-973f0add78fd"
    item-Path="/DemoDocs/AdminDemo"
  ></mgt-file>
`;

export const getFileByUserId = () => html`
  <mgt-person user-id="e3d0513b-449e-4198-ba6f-bd97ae7cae85" view="twolines"></mgt-person>
  <mgt-file user-id="e3d0513b-449e-4198-ba6f-bd97ae7cae85" item-id="01HT2SVWU2EX3RVDWYAFAZDLCGYFZTCUHC"></mgt-file>
  <mgt-file user-id="e3d0513b-449e-4198-ba6f-bd97ae7cae85" item-path="Fashion Products v2.pdf"></mgt-file>
`;

export const getFileByInsights = () => html`
  <mgt-file
    insight-type="shared"
    insight-id="AW1GxMvkOztMkJX-SCppUSRPF5EvyPDHRZVAqtQZXI4JoUXwI9_mfEWNlabr1f7aXRBWDMt2C2FDop4fP1vsUw9tRsTL5Ds7TJCV_kgqaVEkBA"
  ></mgt-file>
`;

export const darkTheme = () => html`
  <mgt-file file-query="/me/drive/items/01BYE5RZZFWGWWVNHHKVHYXE3OUJHGWCT2" class="mgt-dark"></mgt-file>
`;
