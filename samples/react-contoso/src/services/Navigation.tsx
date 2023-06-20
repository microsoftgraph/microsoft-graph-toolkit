import { NavigationItem } from '../models/NavigationItem';
import {
  HomeRegular,
  SearchRegular,
  TextBulletListSquareRegular,
  CalendarMailRegular,
  DocumentRegular
  //ChatRegular
} from '@fluentui/react-icons';
import { DashboardPage } from '../pages/DashboardPage';
import { OutlookPage } from '../pages/OutlookPage';
import { SearchPage } from '../pages/SearchPage';
import { HomePage } from '../pages/HomePage';
import { FilesPage } from '../pages/FilesPage';
//import { ChatPage } from '../pages/ChatPage';

export const getNavigation = (isSignedIn: boolean) => {
  let navItems: NavigationItem[] = [];

  navItems.push({
    name: 'Home',
    url: '/',
    icon: <HomeRegular />,
    key: 'home',
    requiresLogin: false,
    component: <HomePage />,
    exact: true
  });

  if (isSignedIn) {
    navItems.push({
      name: 'Dashboard',
      url: '/dashboard',
      icon: <TextBulletListSquareRegular />,
      key: 'dashboard',
      requiresLogin: true,
      component: <DashboardPage />,
      exact: true
    });

    navItems.push({
      name: 'Mail and Calendar',
      url: '/outlook',
      icon: <CalendarMailRegular />,
      key: 'outlook',
      requiresLogin: true,
      component: <OutlookPage />,
      exact: true
    });

    navItems.push({
      name: 'Files',
      url: '/files',
      icon: <DocumentRegular />,
      key: 'files',
      requiresLogin: true,
      component: <FilesPage />,
      exact: true
    });

    /*navItems.push({
      name: 'Chat',
      url: '/chat',
      icon: <ChatRegular />,
      key: 'chat',
      requiresLogin: true,
      component: <ChatPage />,
      exact: true
    });*/

    navItems.push({
      name: 'Search',
      url: '/search',
      pattern: '/search/:query',
      icon: <SearchRegular />,
      key: 'search',
      requiresLogin: true,
      component: <SearchPage />,
      exact: false
    });
  }
  return navItems;
};
