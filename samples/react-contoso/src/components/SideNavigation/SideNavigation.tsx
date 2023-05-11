import { Tab, TabList, TabValue, makeStyles } from '@fluentui/react-components';
import * as React from 'react';
import { useHistory } from 'react-router-dom';
import { NavigationRegular } from '@fluentui/react-icons';
import { useAppContext } from '../../hooks/useAppContext';

export interface ISideNavigationProps {
  items: any[];
}

const useStyles = makeStyles({
  tab: {
    paddingTop: '12px',
    paddingBottom: '12px'
  },
  activeTab: {
    backgroundColor: 'var(--colorSubtleBackgroundHover)'
  }
});

export const SideNavigation: React.FunctionComponent<ISideNavigationProps> = props => {
  const history = useHistory();
  const { items } = props;
  const [selectedTab, setSelectedTab] = React.useState<TabValue>('/home');
  const [isMinimized, setIsMinimized] = React.useState<boolean>(false);
  const styles = useStyles();
  const appContext = useAppContext();

  const navigate = (event: any, data: any) => {
    if (data.value === 'navigation') {
      const futureIsMinimized = !isMinimized;
      setIsMinimized(futureIsMinimized);
      appContext.setState({ ...appContext.state, isSideBarMinimized: futureIsMinimized });
    } else {
      setSelectedTab(data.value);
      history.push(data.value);
    }
  };

  React.useEffect(() => {
    const pathname = history.location.pathname;
    setSelectedTab(pathname);
  }, [setSelectedTab, history.location.pathname]);

  return (
    <>
      <TabList size="medium" appearance="subtle" vertical onTabSelect={navigate} selectedValue={selectedTab}>
        <Tab icon={<NavigationRegular />} value={'navigation'} className={styles.tab}></Tab>
        {items.map((item, index) => (
          <Tab
            icon={item.icon}
            value={item.url}
            key={index}
            className={`${styles.tab} ${item.url === selectedTab ? styles.activeTab : ''}`}
          >
            {!isMinimized ? item.name : ''}
          </Tab>
        ))}
      </TabList>
    </>
  );
};
