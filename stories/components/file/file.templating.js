import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-file / Templating',
  component: 'file',
  decorators: [withCodeEditor]
};

export const defaultTemplates = () => html`
  <mgt-file file-query="/me/drive/items/01BYE5RZZFWGWWVNHHKVHYXE3OUJHGWCT2">
    <template data-type="default">
      <div class="root">
       <span>Found the file {{file.name}}</span>
       <p>The templateRendered event has been fired!</p>
      </div>
    </template>
    <template data-type="loading">
      <div class="root">
        loading file
      </div>
    </template>
    <template data-type="no-data">
      <div class="root">
        there is no data
      </div>
    </template>
  </mgt-file>
`;

export const templateContext = () => html`
   <mgt-file file-query="/me/drive/items/01BYE5RZZFWGWWVNHHKVHYXE3OUJHGWCT2">
    <template data-type="default">
      <div class="root">
       <span>Found the file {{file.name}} created on {{dayFromDateTime(file.createdDateTime)}}
      </div>
    </template>
    <template data-type="loading">
      <div class="root">
        loading file
      </div>
    </template>
    <template data-type="no-data">
      <div class="root">
        there is no data
      </div>
    </template>
  </mgt-file>
   <script>
     const file = document.querySelector('mgt-file').templateContext = {
      dayFromDateTime: dateTimeString => {
        let date = new Date(dateTimeString);
        date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
        let monthNames = [
          'January',
          'February',
          'March',
          'April',
          'May',
          'June',
          'July',
          'August',
          'September',
          'October',
          'November',
          'December'
        ];

        let monthIndex = date.getMonth();
        let day = date.getDate();
        let year = date.getFullYear();

        return monthNames[monthIndex] + ' ' + day + ' ' + year;
      }
      }
   </script>
 `;
