import React from 'react';
import { HashRouter, Redirect, Route, Switch } from 'react-router-dom';
import './App.css';
import { INavLinkGroup, IStackProps, mergeStyles, IStackTokens, Stack, INavLink } from '@fluentui/react';
import { Header } from './components/Header';
import { SideNavigation } from './components/SideNavigation/SideNavigation';
import { Home } from './pages/Home';
import { SearchCenter } from './pages/SearchCenter';
import { useIsSignedIn } from './hooks/useIsSignedIn';

export const App: React.FunctionComponent = theme => {
  const [navigationItems, setNavigationItems] = React.useState<INavLinkGroup[]>([]);

  const [isSignedIn] = useIsSignedIn();
  const sectionStackTokens: IStackTokens = { childrenGap: 10 };
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
    <Stack horizontal tokens={sectionStackTokens} className={mergeStyles({ overflow: 'hidden' })} {...props} />
  );

  const Sidebar = (props: IStackProps) => (
    <Stack
      disableShrink
      tokens={sidebarStackTokens}
      grow={20}
      className={mergeStyles({
        position: 'fixed',
        top: 48,
        left: 0,
        backgroundColor: 'rgb(233, 233, 233)'
      })}
      verticalFill={true}
      {...props}
    />
  );

  const Main = (props: IStackProps) => (
    <Stack
      grow={80}
      className={mergeStyles({ padding: '40px', paddingLeft: '300px', paddingTop: '42px' })}
      disableShrink
      {...props}
    />
  );

  const Page = (props: IStackProps) => (
    <Stack
      tokens={sectionStackTokens}
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
              {isSignedIn && <Route exact path="/search" component={SearchCenter} />}
              <Route path="*" component={Home} />
            </Switch>
          </Main>
        </HashRouter>
      </Content>
    </Page>
  );
};
