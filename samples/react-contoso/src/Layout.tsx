import React from 'react';
import { HashRouter, Route, Switch } from 'react-router-dom';
import { Header } from './components/Header';
import { SideNavigation } from './components/SideNavigation';
import { HomePage } from './pages/HomePage';
import { useIsSignedIn } from './hooks/useIsSignedIn';
import { NavigationItem } from './models/NavigationItem';
import { getNavigation } from './services/Navigation';
import { FluentProvider, makeStyles, mergeClasses, shorthands } from '@fluentui/react-components';
import { tokens } from '@fluentui/react-theme';
import { applyTheme } from '@microsoft/mgt-react';
import { useAppContext } from './AppContext';

const useStyles = makeStyles({
  sidebar: {
    display: 'flex',
    flexDirection: 'column',
    flexWrap: 'nowrap',
    height: '100%',
    minWidth: '295px',
    boxSizing: 'border-box',
    backgroundColor: tokens.colorNeutralBackground6
  },
  main: {
    backgroundColor: tokens.colorNeutralBackground1,
    display: 'flex',
    flexDirection: 'row',
    width: 'auto',
    height: 'calc(100vh - 50px)',
    boxSizing: 'border-box'
  },
  minimized: {
    minWidth: 'auto'
  },
  page: {
    display: 'flex',
    flexDirection: 'column',
    flexWrap: 'nowrap'
  },
  content: {
    display: 'flex',
    flexDirection: 'column',
    flexWrap: 'nowrap',
    width: '100%',
    height: 'auto',
    boxSizing: 'border-box',
    ...shorthands.margin('10px'),
    ...shorthands.overflow('auto')
  }
});

export const Layout: React.FunctionComponent = theme => {
  const styles = useStyles();
  const [navigationItems, setNavigationItems] = React.useState<NavigationItem[]>([]);
  const [isSignedIn] = useIsSignedIn();
  const appContext = useAppContext();

  React.useEffect(() => {
    setNavigationItems(getNavigation(isSignedIn));
  }, [isSignedIn]);

  React.useEffect(() => {
    // Applies the theme to the MGT components
    applyTheme(appContext.state.theme.key as any);
  }, [appContext]);

  return (
    <FluentProvider theme={appContext.state.theme.fluentTheme}>
      <div className={styles.page}>
        <HashRouter>
          <Header></Header>
          <div className={styles.main}>
            <div
              className={mergeClasses(
                styles.sidebar,
                `${appContext.state.sidebar.isMinimized ? styles.minimized : ''}`
              )}
            >
              <SideNavigation items={navigationItems}></SideNavigation>
            </div>
            <div className={styles.content}>
              <Switch>
                {navigationItems.map(
                  item =>
                    ((item.requiresLogin && isSignedIn) || !item.requiresLogin) && (
                      <Route exact={item.exact} path={item.url} children={item.component} key={item.key} />
                    )
                )}
                <Route path="*" component={HomePage} />
              </Switch>
            </div>
          </div>
        </HashRouter>
      </div>
    </FluentProvider>
  );
};
