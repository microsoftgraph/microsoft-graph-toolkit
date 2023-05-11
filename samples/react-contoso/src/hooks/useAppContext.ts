import React from 'react';
import { AppContext } from '../App';

export function useAppContext() {
  const value = React.useContext(AppContext);
  if (value === undefined) throw new Error('Expected an AppProvider somewhere in the react tree to set context value');
  return value; // now has type AppContextValue
  // or even provide domain methods for better encapsulation
}
