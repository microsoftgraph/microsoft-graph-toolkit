/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

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
export const lazyLoadComponent = <P>(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  component: React.LazyExoticComponent<any>,
  props: P
): React.FunctionComponentElement<React.SuspenseProps> => {
  // imperative code analogous to:
  // <div><Suspense fallback="Loading..."><Component {...props} /></Suspense></div>
  return React.createElement(
    React.Suspense,
    {
      fallback: React.createElement('div', null, 'Loading...')
    },
    React.createElement(component, props)
  );
};
