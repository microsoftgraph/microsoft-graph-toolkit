/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-agenda / React',
  component: 'agenda',
  decorators: [withCodeEditor]
};

export const Agenda = () => html`
  <mgt-agenda></mgt-agenda>
  <react>
    import { Agenda } from '@microsoft/mgt-react';

    export default () => (
      <Agenda></Agenda>
    );
  </react>
`;

export const getByEventQuery = () => html`
  <mgt-agenda event-query="/me/calendarview?$orderby=start/dateTime&startdatetime=2023-07-12T00:00:00.000Z&enddatetime=2023-07-18T00:00:00.000Z"></mgt-agenda>
<react>
import { Agenda } from '@microsoft/mgt-react';

export default () => (
  <Agenda eventQuery='/me/calendarview?$orderby=start/dateTime&startdatetime=2023-07-12T00:00:00.000Z&enddatetime=2023-07-18T00:00:00.000Z'></Agenda>
);
</react>
`;

export const groupByDay = () => html`
  <mgt-agenda group-by-day></mgt-agenda>
<react>
import { Agenda } from '@microsoft/mgt-react';

export default () => (
  <Agenda groupByDay={true}></Agenda>
);
</react>
`;

export const showMax = () => html`
  <mgt-agenda show-max="4"></mgt-agenda>
  <react>
import { Agenda } from '@microsoft/mgt-react';

export default () => (
  <Agenda showMax={4}></Agenda>
);
</react>
`;

export const getByDate = () => html`
  <mgt-agenda group-by-day date="May 7, 2019" days="3"></mgt-agenda>
<react>
import { Agenda } from '@microsoft/mgt-react';

export default () => (
  <Agenda groupByDay={true} date='May 7, 2019' days={3}></Agenda>
);
</react>
`;

export const getEventsForNextWeek = () => html`
  <mgt-agenda group-by-day days="7"></mgt-agenda>
<react>
import { Agenda } from '@microsoft/mgt-react';

export default () => (
  <Agenda groupByDay={true} days={7}></Agenda>
);
</react>
`;

export const preferredTimezone = () => html`
  <mgt-agenda preferred-timezone="Europe/Paris"></mgt-agenda>
<react>
import { Agenda } from '@microsoft/mgt-react';

export default () => (
  <Agenda preferredTimezone='Europe/Paris'></Agenda>
);
</react>
`;

export const events = () => html`
  <mgt-agenda></mgt-agenda>
  <react>
    import { Agenda } from '@microsoft/mgt-react';

    export default () => (
      const onUpdated = (e) => {
        console.log('updated', e);
      };

      const onEventClick = (e) => {
        console.log(e.detail);
      };

      return(
        <Agenda
          updated={onUpdated}
          eventClick={onEventClick}>
        </Agenda>
      );
    );
  </react>
  <script>
    document.querySelector('mgt-agenda').addEventListener('updated', e => {
      console.log('updated', e);
    });
    document.querySelector('mgt-agenda').addEventListener('eventClick', e => {
      console.log(e.detail);
    });
  </script>
`;
