import * as React from 'react';
import { PageHeader } from '../components/PageHeader';
import { FileList } from '@microsoft/mgt-react';
import {
  SelectTabData,
  SelectTabEvent,
  Tab,
  TabList,
  TabValue,
  shorthands,
  makeStyles
} from '@fluentui/react-components';
import { ChannelFiles } from './Files/ChannelFiles';
import { SiteFiles } from './Files/SiteFiles';

const useStyles = makeStyles({
  panels: {
    ...shorthands.padding('10px')
  }
});

const FilesPage: React.FunctionComponent = () => {
  const styles = useStyles();
  const [selectedTab, setSelectedTab] = React.useState<TabValue>('my');

  const onTabSelect = (event: SelectTabEvent, data: SelectTabData) => {
    setSelectedTab(data.value);
  };

  return (
    <>
      <PageHeader
        title={'Files'}
        description={
          'View your files from accross your OneDrive, channels you are a member of and your SharePoint sites'
        }
      ></PageHeader>

      <div>
        <TabList selectedValue={selectedTab} onTabSelect={onTabSelect}>
          <Tab value="my">My Files</Tab>
          <Tab value="recent">Recent Files</Tab>
          <Tab value="site">Site Files</Tab>
          <Tab value="channel">Channel Files</Tab>
        </TabList>
        <div className={styles.panels}>
          {selectedTab === 'my' && <FileList pageSize={100}></FileList>}
          {selectedTab === 'recent' && <FileList insightType="used" enableFileUpload={false} pageSize={100}></FileList>}
          {selectedTab === 'site' && <SiteFiles />}
          {selectedTab === 'channel' && <ChannelFiles />}
        </div>
      </div>
    </>
  );
};

export default FilesPage;
