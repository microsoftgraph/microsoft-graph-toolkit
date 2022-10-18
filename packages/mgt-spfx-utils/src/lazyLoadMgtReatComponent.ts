import * as React from 'react';
export function lazyLoadComponent<P>(
  component: React.LazyExoticComponent<any>,
  props: P
): React.FunctionComponentElement<React.SuspenseProps> {
  return React.createElement(
    React.Suspense,
    {
      fallback: React.createElement('div', null, 'Loading...')
    },
    React.createElement(component, props)
  );
}
