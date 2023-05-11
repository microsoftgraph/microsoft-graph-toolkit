import * as React from 'react';
import { PageHeader } from '../components/PageHeader/PageHeader';
import { Get } from '@microsoft/mgt-react';
import { Messages } from '../components/Messages/Messages';
import { Loading } from '../components/Loading';
import {
  SelectTabData,
  SelectTabEvent,
  Tab,
  TabList,
  TabValue,
  shorthands,
  makeStyles
} from '@fluentui/react-components';

const useStyles = makeStyles({
  panels: {
    ...shorthands.padding('10px')
  }
});

export const Emails: React.FunctionComponent = () => {
  const styles = useStyles();
  const [selectedTab, setSelectedTab] = React.useState<TabValue>('focused');

  const onTabSelect = (event: SelectTabEvent, data: SelectTabData) => {
    setSelectedTab(data.value);
  };

  return (
    <>
      <PageHeader title={'My Emails'} description={'Your Microsoft 365 emails.'}></PageHeader>

      <div>
        <TabList selectedValue={selectedTab} onTabSelect={onTabSelect}>
          <Tab value="focused">Focused</Tab>
          <Tab value="others">Others</Tab>
        </TabList>
        <div className={styles.panels}>
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
      </div>
    </>
  );
};
