import * as React from 'react';
import { PageHeader } from '../components/PageHeader';
import { Todo } from '@microsoft/mgt-react';
import { DirectReports } from '../components/DirectReports';
import {
  SelectTabData,
  SelectTabEvent,
  Tab,
  TabList,
  TabValue,
  makeStyles,
  shorthands
} from '@fluentui/react-components';

const useStyles = makeStyles({
  panels: {
    ...shorthands.padding('10px')
  }
});

const DashboardPage: React.FunctionComponent = () => {
  const styles = useStyles();

  const [selectedTab, setSelectedTab] = React.useState<TabValue>('tasks');

  const onTabSelect = (event: SelectTabEvent, data: SelectTabData) => {
    setSelectedTab(data.value);
  };

  return (
    <>
      <PageHeader
        title={'My Dashboard'}
        description={'This dashboard helps you being productive with your tasks and your incidents.'}
      ></PageHeader>
      <div>
        <TabList selectedValue={selectedTab} onTabSelect={onTabSelect}>
          <Tab value="tasks">My Tasks</Tab>
          <Tab value="directReports">My Direct Reports</Tab>
        </TabList>
        <div className={styles.panels}>
          {selectedTab === 'tasks' && <Todo></Todo>}
          {selectedTab === 'directReports' && <DirectReports />}
        </div>
      </div>
    </>
  );
};

export default DashboardPage;
