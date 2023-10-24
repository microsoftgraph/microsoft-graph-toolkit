import * as React from 'react';
import { PageHeader } from '../components/PageHeader';
import { Agenda, Get } from '@microsoft/mgt-react';
import { Messages } from '../components/Messages';
import { Loading } from '../components/Loading';
import {
  SelectTabData,
  SelectTabEvent,
  Tab,
  TabList,
  TabValue,
  shorthands,
  makeStyles,
  mergeClasses
} from '@fluentui/react-components';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'row'
  },
  panels: {
    ...shorthands.padding('10px')
  },
  main: {
    display: 'flex',
    flexDirection: 'column',
    flexWrap: 'nowrap',
    width: '70%'
  },
  side: {
    display: 'flex',
    flexDirection: 'column',
    flexWrap: 'nowrap',
    width: '30%'
  }
});

const OutlookPage: React.FunctionComponent = () => {
  const styles = useStyles();
  const [selectedTab, setSelectedTab] = React.useState<TabValue>('focused');

  const onTabSelect = (event: SelectTabEvent, data: SelectTabData) => {
    setSelectedTab(data.value);
  };

  return (
    <>
      <PageHeader
        title={'Mail and Calendar'}
        description={'Stay productive and navigate your emails and your calendar appointments'}
      ></PageHeader>

      <TabList selectedValue={selectedTab} onTabSelect={onTabSelect} className={styles.container}>
        <Tab value="focused">Focused</Tab>
        <Tab value="others">Others</Tab>
      </TabList>
      <div className={styles.container}>
        <div className={mergeClasses(styles.panels, styles.main)}>
          {selectedTab === 'focused' && (
            <Get
              resource="/me/mailFolders/Inbox/messages?$orderby=InferenceClassification, createdDateTime DESC&filter=InferenceClassification eq 'Focused'"
              maxPages={3}
              scopes={['Mail.Read']}
            >
              <Messages template="value"></Messages>
              <Loading template="loading" message={'Loading your focused inbox...'}></Loading>
            </Get>
          )}
          {selectedTab === 'others' && (
            <Get
              resource="/me/mailFolders/Inbox/messages?$orderby=InferenceClassification, createdDateTime DESC&filter=InferenceClassification eq 'Other'"
              maxPages={3}
              scopes={['Mail.Read']}
            >
              <Messages template="value"></Messages>
              <Loading template="loading" message={'Loading your other inbox...'}></Loading>
            </Get>
          )}
        </div>
        <div className={styles.side}>
          <Agenda groupByDay={true}></Agenda>
        </div>
      </div>
    </>
  );
};

export default OutlookPage;
