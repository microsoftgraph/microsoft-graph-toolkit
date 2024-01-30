import * as React from 'react';
import { PageHeader } from '../components/PageHeader';
import {
  FileList,
  File,
  Person,
  Spinner,
  Agenda,
  PersonCard,
  People,
  Get,
  ThemeToggle,
  Picker,
  TeamsChannelPicker
} from '@microsoft/mgt-react';
import { makeStyles } from '@fluentui/react-components';
import { Messages } from '../components/Messages';
import { Loading } from '../components/Loading';
import { GetDefaultContent } from '../components/GetDefaultContent';
import { Stylesheet } from '@fluentui/merge-styles';

const useStyles = makeStyles({
  root: {
    '&::part(detail-line)': {
      whiteSpace: 'nowrap',
      textOverflow: 'ellipsis',
      width: '-webkit-fill-available',
      overflowX: 'hidden',
      overflowY: 
    }
  },
  container: {
    width: '150px'
  }
});

const HomePage: React.FunctionComponent = () => {
  const styles = useStyles();
  return (
    <>
      {/* <ThemeToggle /> */}
      {/* <TeamsChannelPicker /> */}
      {/* <Picker
        resource="me/todo/lists"
        scopes={['tasks.read', 'tasks.readwrite']}
        placeholder="Select a task list"
        keyName="displayName"
      />
      <FileList />
      <Get resource="/me/messages" maxPages={1}>
        <GetDefaultContent template="default" />
        <Messages template="value"></Messages>
        <Loading template="loading" message={'Loading your focused inbox...'}></Loading>
      </Get> */}
      <PageHeader title={'Home'} description={'Welcome to Contoso!'}></PageHeader>
      {/* <Spinner />
      <People />
      <File
        driveId="b!M5IeZ2QKf0y18TIIXsDQkecHx1QrukxCte8X3n6ka6yn409-utaER7M2W9uRO4yB"
        itemId="01WEUQSTSBWERA5VH4BFALQBXUDVUMT22G"
      /> */}
      {/* <Person
        personQuery="LeeG@wgww6.onmicrosoft.com"
        personCardInteraction="click"
        avatarType="initials"
        avatarSize="auto"
      /> */}
      {/* <div className={styles.container}>
        <Person className={styles.root} personQuery="me" view="threelines" line2Property="email" />
      </div> */}
      {/* <Agenda /> */}
      {/* <PersonCard personQuery="LeeG@wgww6.onmicrosoft.com" /> */}
    </>
  );
};

export default HomePage;
