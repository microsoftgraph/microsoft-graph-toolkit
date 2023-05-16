import * as React from 'react';
import { PageHeader } from '../components/PageHeader/PageHeader';
import { AllResults, PeopleResults } from './Search';
import {
  SelectTabData,
  SelectTabEvent,
  Tab,
  TabList,
  TabValue,
  makeStyles,
  shorthands
} from '@fluentui/react-components';
import { useAppContext } from '../AppContext';

const useStyles = makeStyles({
  panels: {
    ...shorthands.padding('10px')
  },
  container: {
    maxWidth: '1028px',
    width: '100%'
  }
});

export const SearchPage: React.FunctionComponent = () => {
  const styles = useStyles();
  const appContext = useAppContext();

  const [selectedTab, setSelectedTab] = React.useState<TabValue>('allResults');

  const onTabSelect = (event: SelectTabEvent, data: SelectTabData) => {
    setSelectedTab(data.value);
  };

  return (
    <>
      <PageHeader
        title={'Search'}
        description={'Use this Search Center to test Microsot Graph Toolkit search components capabilities'}
      ></PageHeader>

      <div className={styles.container}>
        <TabList selectedValue={selectedTab} onTabSelect={onTabSelect}>
          <Tab value="allResults">All Results</Tab>
          <Tab value="people">People</Tab>
        </TabList>
        <div className={styles.panels}>
          {selectedTab === 'allResults' && <AllResults searchTerm={appContext.state.searchTerm}></AllResults>}
          {selectedTab === 'people' && <PeopleResults searchTerm={appContext.state.searchTerm}></PeopleResults>}
        </div>
      </div>
    </>
  );
};
