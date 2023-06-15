import { Tab, TabList, makeStyles, mergeClasses, tokens } from '@fluentui/react-components';
import * as React from 'react';
import { useHistory, useLocation } from 'react-router-dom';
import { NavigationRegular } from '@fluentui/react-icons';
import { useAppContext } from '../AppContext';
import { NavigationItem } from '../models/NavigationItem';

export interface ISideNavigationProps {
  items: any[];
}

const useStyles = makeStyles({
  tab: {
    paddingTop: '12px',
    paddingBottom: '12px'
  },
  activeTab: {
    backgroundColor: tokens.colorSubtleBackgroundHover
  }
});

export const SideNavigation: React.FunctionComponent<ISideNavigationProps> = props => {
  const history = useHistory();
  const location = useLocation();
  const { items } = props;
  const [selectedTab, setSelectedTab] = React.useState<any>('/');
  const [isMinimized, setIsMinimized] = React.useState<boolean>(false);
  const styles = useStyles();
  const appContext = useAppContext();

  const navigate = (event: any, data: any) => {
    if (data.value === 'navigation') {
      const futureIsMinimized = !isMinimized;
      setIsMinimized(futureIsMinimized);
      appContext.setState({
        ...appContext.state,
        sidebar: { ...appContext.state.sidebar, isMinimized: futureIsMinimized }
      });
    } else {
      setSelectedTab(data.value);
      history.push(data.value);
    }
  };

  React.useLayoutEffect(() => {
    setSelectedTab(location.pathname);
  }, [location]);

  return (
    <>
      <TabList size="medium" appearance="subtle" vertical onTabSelect={navigate} selectedValue={selectedTab}>
        <Tab icon={<NavigationRegular />} value={'navigation'} className={styles.tab}></Tab>
        {items.map((item: NavigationItem, index) => (
          <Tab
            icon={item.icon}
            value={item.url}
            key={index}
            className={mergeClasses(styles.tab, item.url === selectedTab && styles.activeTab)}
          >
            {!isMinimized ? item.name : ''}
          </Tab>
        ))}
      </TabList>
    </>
  );
};
