import * as React from 'react';
import { PageHeader } from '../components/PageHeader';
import { Providers, Todo } from '@microsoft/mgt-react';
import { Incidents } from '../components/Incidents';
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

export const DashboardPage: React.FunctionComponent = () => {
  const styles = useStyles();

  const [taskListId, setTaskListId] = React.useState<string>('');
  const [selectedTab, setSelectedTab] = React.useState<TabValue>('tasks');

  const onTabSelect = (event: SelectTabEvent, data: SelectTabData) => {
    setSelectedTab(data.value);
  };

  React.useEffect(() => {
    const fetchTaskList = async () => {
      const taskList = await Providers.globalProvider.graph.client.api('/me/todo/lists?$top=1').get();
      setTaskListId(taskList.value[0].id);
    };

    fetchTaskList();
  }, []);

  return (
    <>
      <PageHeader
        title={'My Dashboard'}
        description={'This dashboard helps you being productive with your tasks and your incidents.'}
      ></PageHeader>
      <div>
        <TabList selectedValue={selectedTab} onTabSelect={onTabSelect}>
          <Tab value="tasks">My Tasks</Tab>
          <Tab value="incidents">My Incidents</Tab>
        </TabList>
        <div className={styles.panels}>
          {selectedTab === 'tasks' && taskListId && <Todo targetId={taskListId}></Todo>}
          {selectedTab === 'incidents' && <Incidents />}
        </div>
      </div>
    </>
  );
};
