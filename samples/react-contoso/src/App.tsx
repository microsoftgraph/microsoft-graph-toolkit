import React, { Dispatch, SetStateAction } from 'react';
import { HashRouter, Redirect, Route, Switch } from 'react-router-dom';
import { Header } from './components/Header';
import { SideNavigation } from './components/SideNavigation/SideNavigation';
import { Home } from './pages/Home';
import { SearchCenter } from './pages/SearchCenter';
import { useIsSignedIn } from './hooks/useIsSignedIn';
import { Emails } from './pages/Emails';
import { Dashboard } from './pages/Dashboard';
import { Incident } from './pages/Incident/Incident';
import './App.css';
import { HomeRegular, SearchRegular, MailRegular, TextBulletListSquareRegular } from '@fluentui/react-icons';

type AppContextState = { searchTerm: string; isSideBarMinimized: boolean };

type AppContextValue = {
  state: AppContextState;
  setState: Dispatch<SetStateAction<AppContextState>>;
};

export const AppContext = React.createContext<AppContextValue | undefined>(undefined);

export const App: React.FunctionComponent = theme => {
  const [navigationItems, setNavigationItems] = React.useState<any[]>([]);
  const [isSignedIn] = useIsSignedIn();
  const [state, setState] = React.useState({ searchTerm: '', isSideBarMinimized: false });

  React.useEffect(() => {
    let navItems: any[] = [];

    navItems.push({
      name: 'Home',
      url: '/home',
      icon: <HomeRegular />,
      key: 'home',
      requiresLogin: false
    });

    if (isSignedIn) {
      navItems.push({
        name: 'Dashboard',
        url: '/dashboard',
        icon: <TextBulletListSquareRegular />,
        key: 'dashboard',
        requiresLogin: true
      });

      navItems.push({
        name: 'Messages',
        url: '/messages',
        icon: <MailRegular />,
        key: 'messages',
        requiresLogin: true
      });

      navItems.push({
        name: 'Search',
        url: '/search',
        icon: <SearchRegular />,
        key: 'search',
        requiresLogin: true
      });
    }

    setNavigationItems(navItems);
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
                <Route exact path="/home" component={Home} />
                {isSignedIn && <Route exact path="/dashboard" component={Dashboard} />}
                {isSignedIn && <Route exact path="/search" component={SearchCenter} />}
                {isSignedIn && <Route exact path="/messages" component={Emails} />}
                {isSignedIn && <Route path="/incident/:id" component={Incident} />}
                <Route path="*" component={Home} />
              </Switch>
            </div>
          </div>
        </HashRouter>
      </AppContext.Provider>
    </div>
  );
};
