import { lazy, Suspense } from 'react';
import { NavigationItem } from '../models/NavigationItem';
import {
  HomeRegular,
  SearchRegular,
  TextBulletListSquareRegular,
  CalendarMailRegular,
  DocumentRegular,
  TagMultipleRegular
} from '@fluentui/react-icons';
const DashboardPage = lazy(() => import('../pages/DashboardPage'));
const OutlookPage = lazy(() => import('../pages/OutlookPage'));
const SearchPage = lazy(() => import('../pages/SearchPage'));
const HomePage = lazy(() => import('../pages/HomePage'));
const FilesPage = lazy(() => import('../pages/FilesPage'));
const TaxonomyPage = lazy(() => import('../pages/TaxonomyPage'));

export const getNavigation = (isSignedIn: boolean) => {
  let navItems: NavigationItem[] = [];

  navItems.push({
    name: 'Home',
    url: '',
    icon: <HomeRegular />,
    key: 'home',
    requiresLogin: false,
    component: (
      <Suspense fallback="Loading...">
        <HomePage />
      </Suspense>
    ),
    exact: true
  });

  if (isSignedIn) {
    navItems.push({
      name: 'Dashboard',
      url: 'dashboard',
      icon: <TextBulletListSquareRegular />,
      key: 'dashboard',
      requiresLogin: true,
      component: (
        <Suspense fallback="Loading...">
          <DashboardPage />
        </Suspense>
      ),
      exact: true
    });

    navItems.push({
      name: 'Mail and Calendar',
      url: 'outlook',
      icon: <CalendarMailRegular />,
      key: 'outlook',
      requiresLogin: true,
      component: (
        <Suspense fallback="Loading...">
          <OutlookPage />
        </Suspense>
      ),
      exact: true
    });

    navItems.push({
      name: 'Files',
      url: 'files',
      icon: <DocumentRegular />,
      key: 'files',
      requiresLogin: true,
      component: (
        <Suspense fallback="Loading...">
          <FilesPage />
        </Suspense>
      ),
      exact: true
    });

    navItems.push({
      name: 'Taxonomy',
      url: 'taxonomy',
      icon: <TagMultipleRegular />,
      key: 'files',
      requiresLogin: true,
      component: (
        <Suspense fallback="Loading...">
          <TaxonomyPage />
        </Suspense>
      ),
      exact: true
    });

    navItems.push({
      name: 'Search',
      url: 'search',
      pattern: '/search/:query',
      icon: <SearchRegular />,
      key: 'search',
      requiresLogin: true,
      component: (
        <Suspense fallback="Loading...">
          <SearchPage />
        </Suspense>
      ),
      exact: false
    });
  }
  return navItems;
};
