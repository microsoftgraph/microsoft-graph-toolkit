import * as React from 'react';
import { PageHeader } from '../components/PageHeader';

const HomePage: React.FunctionComponent = () => {
  return (
    <>
      <PageHeader title={'Home'} description={'Welcome to Contoso!'}></PageHeader>
    </>
  );
};

export default HomePage;
