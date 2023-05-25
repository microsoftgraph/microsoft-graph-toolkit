import React from 'react';
import { BrowserRouter, Route, Switch } from 'react-router-dom';
import { Header } from './components/Header';
import { SideNavigation } from './components/SideNavigation/SideNavigation';
import { HomePage } from './pages/HomePage';
import { useIsSignedIn } from './hooks/useIsSignedIn';
import { Incident } from './pages/Incident/Incident';
import './App.css';
import { NavigationItem } from './models/NavigationItem';
import { getNavigation } from './services/Navigation';
import { FluentProvider, makeStyles, mergeClasses } from '@fluentui/react-components';
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
  minimized: {
    minWidth: 'auto'
  },
  mgtFluent: {
    /*'--accent-fill-rest': tokens.colorCompoundBrandForeground1,
    '--accent-fill-active': tokens.colorCompoundBrandForeground1Pressed,
    '--accent-fill-hover': tokens.colorCompoundBrandForeground1Hover,
    '--accent-foreground-rest': tokens.colorCompoundBrandForeground1,
    '--accent-foreground-hover': tokens.colorCompoundBrandForeground1Hover,
    '--accent-stroke-control-rest': tokens.colorCompoundBrandStroke,
    '--accent-stroke-control-active': tokens.colorCompoundBrandStrokePressed,
    '--accent-base-color': tokens.colorNeutralForeground1,
    '--input-background-color': tokens.colorNeutralBackground1*/
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
      <div className={mergeClasses('page', appContext.state.theme.key === 'dark' ? styles.mgtFluent : '')}>
        <BrowserRouter>
          <Header></Header>
          <div className="main">
            <div
              className={mergeClasses(
                styles.sidebar,
                `${appContext.state.sidebar.isMinimized ? styles.minimized : ''}`
              )}
            >
              <SideNavigation items={navigationItems}></SideNavigation>
            </div>
            <div className="content">
              <Switch>
                {navigationItems.map(
                  item =>
                    ((item.requiresLogin && isSignedIn) || !item.requiresLogin) && (
                      <Route exact path={item.url} children={item.component} key={item.url} />
                    )
                )}

                {isSignedIn && <Route path="/incident/:id" component={Incident} />}
                <Route path="*" component={HomePage} />
              </Switch>
            </div>
          </div>
        </BrowserRouter>
      </div>
    </FluentProvider>
  );
};
