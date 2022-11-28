import * as React from 'react';
/**
 * Function to wrap a lazy load a webpart component inside React.Suspense
 *
 * @export
 * @template P
 * @param {React.LazyExoticComponent<any>} component loaded using React.lazy
 * @param {P} props
 * @return {*}  {React.FunctionComponentElement<React.SuspenseProps>}
 */
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
