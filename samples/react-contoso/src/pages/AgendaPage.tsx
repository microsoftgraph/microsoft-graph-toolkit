import * as React from 'react';
import { PageHeader } from '../components/PageHeader/PageHeader';
import { Agenda } from '@microsoft/mgt-react';

export const AgendaPage: React.FunctionComponent = () => {
  return (
    <>
      <PageHeader title={'My Agenda'} description={'Your Microsoft 365 calendar.'}></PageHeader>

      <div>
        <Agenda></Agenda>
      </div>
    </>
  );
};
