import React from 'react';
import { HashRouter, Redirect, Route, Switch } from 'react-router-dom';
import './App.css';
import { INavLinkGroup, IStackProps, mergeStyles, IStackTokens, Stack, INavLink } from '@fluentui/react';
import { Header } from './components/Header';
import { SideNavigation } from './components/SideNavigation/SideNavigation';
import { Home } from './pages/Home';
import { SearchCenter } from './pages/SearchCenter';
import { useIsSignedIn } from './hooks/useIsSignedIn';
import { Emails } from './pages/Emails';
import { Dashboard } from './pages/Dashboard';
import { Incident } from './pages/Incident/Incident';

export const App: React.FunctionComponent = theme => {
  const [navigationItems, setNavigationItems] = React.useState<INavLinkGroup[]>([]);

  const [isSignedIn] = useIsSignedIn();
  //const sectionStackTokens: IStackTokens = { childrenGap: 10 };
  const sidebarStackTokens: IStackTokens = { childrenGap: 20 };

  React.useEffect(() => {
    let navItems: INavLink[] = [];

    navItems.push({
      name: 'Home',
      url: '#/home',
      icon: 'Home',
      key: 'home',
      requiresLogin: false
    });

    if (isSignedIn) {
      navItems.push({
        name: 'Dashboard',
        url: '#/dashboard',
        icon: 'GoToDashboard',
        key: 'dashboard',
        requiresLogin: true
      });

      navItems.push({
        name: 'Messages',
        url: '#/messages',
        icon: 'Mail',
        key: 'messages',
        requiresLogin: true
      });

      navItems.push({
        name: 'Search',
        url: '#/search',
        icon: 'Search',
        key: 'search',
        requiresLogin: true
      });
    }

    setNavigationItems([
      {
        links: navItems
      }
    ]);
  }, [isSignedIn]);

  const Content = (props: IStackProps) => (
    <Stack horizontal className={mergeStyles({ /*overflow: 'hidden'*/ height: 'calc(100vh - 50px)' })} {...props} />
  );

  const Sidebar = (props: IStackProps) => (
    <Stack
      tokens={sidebarStackTokens}
      className={mergeStyles({
        width: '20%',
        backgroundColor: 'rgb(233, 233, 233)'
      })}
      verticalFill={true}
      {...props}
    />
  );

  const Main = (props: IStackProps) => (
    <Stack
      className={mergeStyles({
        margin: '10px',
        //overflow: 'hidden',
        width: '80%'
        /*padding: '40px',
        paddingLeft: '300px',
        paddingTop: '42px',*/
        //minHeight: 'calc(100vh - 10px)'
      })}
      {...props}
    />
  );

  const Page = (props: IStackProps) => (
    <Stack
      //tokens={sectionStackTokens}
      className={mergeStyles({
        selectors: {
          ':global(body)': {
            padding: 0,
            margin: 0
          }
        }
      })}
      {...props}
    />
  );

  return (
    <Page>
      <Header></Header>
      <Content>
        <HashRouter>
          <Sidebar>
            <SideNavigation items={navigationItems}></SideNavigation>
          </Sidebar>
          <Main>
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
          </Main>
        </HashRouter>
      </Content>
    </Page>
  );
};
