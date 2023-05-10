import * as React from 'react';
import { PageHeader } from '../components/PageHeader/PageHeader';

export const Home: React.FunctionComponent = () => {
  return (
    <PageHeader
      title={'Home'}
      description={'Use this application to test Microsoft Graph Toolkit capabilities in rich scenarios.'}
    ></PageHeader>
  );
};
