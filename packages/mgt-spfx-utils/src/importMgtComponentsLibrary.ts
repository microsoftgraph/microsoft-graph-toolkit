/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * Function to simplify lazy loading @microsoft/mgt-components script
 * Lazy loading of @microsoft/mgt-components is necessary when using
 * the disambiguation feature.
 *
 * @export
 * @param {boolean} hasImportedMgtScripts
 * @param {() => void} onSuccessCallback
 * @param {(e: Error) => void} [onErrorCallback]
 */
export function importMgtComponentsLibrary(
  hasImportedMgtScripts: boolean,
  onSuccessCallback: () => void,
  onErrorCallback?: (e?: Error) => void
) {
  if (!hasImportedMgtScripts) {
    import('@microsoft/mgt-components')
      .then(() => {
        hasImportedMgtScripts = true;
        onSuccessCallback();
      })
      .catch(e => {
        onErrorCallback?.(e);
      });
  }
}
