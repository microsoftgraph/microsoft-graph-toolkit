import React from 'react';
import './App.css';
import { AppContext } from './AppContext';
import { Layout } from './Layout';
import { webLightTheme } from '@fluentui/react-components';
import { lightTheme } from '@microsoft/mgt-chat';

export const App: React.FunctionComponent = theme => {
  const [state, setState] = React.useState({
    searchTerm: '',
    sidebar: {
      isMinimized: false
    },
    theme: { key: 'light', fluentTheme: webLightTheme, chatTheme: lightTheme }
  });

  return (
    <AppContext.Provider value={{ state, setState }}>
      <Layout />
    </AppContext.Provider>
  );
};
