import * as React from 'react';
import { PageHeader } from '../components/PageHeader';
import { People } from '@microsoft/mgt-react';
const peopleDisplay: string[] = ['LidiaH', 'Megan Bowen', 'Lynne Robbins', 'JoniS'];

const HomePage: React.FunctionComponent = () => {
  return (
    <>
      <PageHeader title={'Home'} description={'Welcome to Contoso!'}></PageHeader>
      <People peopleQueries={peopleDisplay}></People>
    </>
  );
};

export default HomePage;
