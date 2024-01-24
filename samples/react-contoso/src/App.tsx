import React from 'react';
import { AppContext } from './AppContext';
import { Layout } from './Layout';
import { webLightTheme } from '@fluentui/react-components';

const App: React.FunctionComponent = theme => {
  const [state, setState] = React.useState({
    searchTerm: '*',
    sidebar: {
      isMinimized: false
    },
    theme: { key: 'light', fluentTheme: webLightTheme }
  });

  return (
    <AppContext.Provider value={{ state, setState }}>
      <Layout />
    </AppContext.Provider>
  );
};

export default App;
