import * as React from 'react';
import { PageHeader } from '../components/PageHeader';
import { Person, PersonViewType, Login, FileList, File, PersonCardInteraction } from '@microsoft/mgt-react';

const HomePage: React.FunctionComponent = () => {
  return (
    <>
      <PageHeader title={'Home'} description={'Welcome to Contoso!'}></PageHeader>
      <Person personQuery="me" view={PersonViewType.twolines}></Person>
      {/* <Login></Login> */}
      {/* <FileList></FileList> */}
      {/* <File fileQuery="/me/drive/items/01SJYAGHXIO2UXFBXQABFLA7ESV7PBSXDP"></File> */}
    </>
  );
};

export default HomePage;
