import * as React from 'react';
import { PageHeader } from '../components/PageHeader';
import { FileList, File, Person, PersonCardInteraction, Spinner, Agenda } from '@microsoft/mgt-react';

const HomePage: React.FunctionComponent = () => {
  return (
    <>
      <PageHeader title={'Home'} description={'Welcome to Contoso!'}></PageHeader>
      <Spinner />
      <File
        driveId="b!M5IeZ2QKf0y18TIIXsDQkecHx1QrukxCte8X3n6ka6yn409-utaER7M2W9uRO4yB"
        itemId="01WEUQSTSBWERA5VH4BFALQBXUDVUMT22G"
      />
      <FileList />
      <Person personQuery="LeeG@wgww6.onmicrosoft.com" personCardInteraction={PersonCardInteraction.click} />
      <Agenda />
    </>
  );
};

export default HomePage;
