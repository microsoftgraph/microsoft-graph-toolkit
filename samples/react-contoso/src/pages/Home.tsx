import * as React from 'react';
import { PageHeader } from '../components/PageHeader/PageHeader';

export const Home: React.FunctionComponent = () => {
  return (
    <>
      <PageHeader title={'Home'} description={'Welcome to Contoso!'}></PageHeader>
    </>
  );
};
