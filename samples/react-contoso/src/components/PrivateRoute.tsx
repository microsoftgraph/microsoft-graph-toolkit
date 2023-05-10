import * as React from 'react';
import { Route, RouteProps } from 'react-router-dom';
import { useIsSignedIn } from '../hooks/useIsSignedIn';
import { Loading } from './Loading';

export interface IPrivateRouteProps extends RouteProps {
  component: any;
}

export const PrivateRoute: React.FunctionComponent<IPrivateRouteProps> = (props: IPrivateRouteProps) => {
  const { component: Component, ...rest } = props;
  const [isSignedIn] = useIsSignedIn();

  return (
    <Route
      {...rest}
      render={routeProps => (isSignedIn ? <Component {...routeProps} /> : <Loading message="Logging in..." />)}
    />
  );
};
