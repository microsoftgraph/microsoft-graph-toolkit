import React, { Dispatch, SetStateAction } from 'react';
import { HashRouter, Redirect, Route, Switch } from 'react-router-dom';
import { Header } from './components/Header';
import { SideNavigation } from './components/SideNavigation/SideNavigation';
import { HomePage } from './pages/HomePage';
import { useIsSignedIn } from './hooks/useIsSignedIn';
import { Incident } from './pages/Incident/Incident';
import './App.css';
import { NavigationItem } from './models/NavigationItem';
import { getNavigation } from './services/Navigation';

type AppContextState = { searchTerm: string; isSideBarMinimized: boolean };

type AppContextValue = {
  state: AppContextState;
  setState: Dispatch<SetStateAction<AppContextState>>;
};

export const AppContext = React.createContext<AppContextValue | undefined>(undefined);

export const App: React.FunctionComponent = theme => {
  const [navigationItems, setNavigationItems] = React.useState<NavigationItem[]>([]);
  const [isSignedIn] = useIsSignedIn();
  const [state, setState] = React.useState({ searchTerm: '', isSideBarMinimized: false });

  React.useEffect(() => {
    setNavigationItems(getNavigation(isSignedIn));
  }, [isSignedIn]);

  return (
    <div className="page">
      <AppContext.Provider value={{ state, setState }}>
        <HashRouter>
          <Header></Header>
          <div className="main">
            <div className={`sidebar ${state.isSideBarMinimized ? 'minimized' : ''}`}>
              <SideNavigation items={navigationItems}></SideNavigation>
            </div>
            <div className="content">
              <Switch>
                <Route exact path="/">
                  <Redirect to="/home" />
                </Route>
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
        </HashRouter>
      </AppContext.Provider>
    </div>
  );
};
