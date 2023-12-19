import * as React from 'react';
import { PageHeader } from '../components/PageHeader';
import {
  FileList,
  File,
  Person,
  PersonCardInteraction,
  Spinner,
  Agenda,
  PersonCard,
  People,
  Get,
  ThemeToggle,
  Picker
} from '@microsoft/mgt-react';
import { Messages } from '../components/Messages';
import { Loading } from '../components/Loading';
import { GetDefaultContent } from '../components/GetDefaultContent';

const HomePage: React.FunctionComponent = () => {
  return (
    <>
      <ThemeToggle />
      <Picker
        resource="me/todo/lists"
        scopes={['tasks.read', 'tasks.readwrite']}
        placeholder="Select a task list"
        keyName="displayName"
      />
      <FileList />
      <Get resource="/me/messages" maxPages={10}>
        <GetDefaultContent template="default" />
        <Messages template="value"></Messages>
        <Loading template="loading" message={'Loading your focused inbox...'}></Loading>
      </Get>
      <PageHeader title={'Home'} description={'Welcome to Contoso!'}></PageHeader>
      <Spinner />
      <People />
      <File
        driveId="b!M5IeZ2QKf0y18TIIXsDQkecHx1QrukxCte8X3n6ka6yn409-utaER7M2W9uRO4yB"
        itemId="01WEUQSTSBWERA5VH4BFALQBXUDVUMT22G"
      />
      <Person personQuery="LeeG@wgww6.onmicrosoft.com" personCardInteraction={PersonCardInteraction.click} />
      <Agenda />
      <PersonCard personQuery="LeeG@wgww6.onmicrosoft.com" />
    </>
  );
};

export default HomePage;
