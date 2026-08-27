/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

let hasWarned = false;

/**
 * Emits a one-time console warning informing consumers that the
 * Microsoft Graph Toolkit has been retired.
 *
 * The warning is emitted at most once per runtime, and any failure to
 * write to the console is swallowed so this can never break a host app.
 */
export const emitDeprecationWarning = (): void => {
  if (hasWarned) {
    return;
  }
  hasWarned = true;

  try {
    console.warn(
      '🦒: The Microsoft Graph Toolkit was retired on August 28, 2026 and is no longer maintained. ' +
        'This package will receive no further updates. ' +
        'If you would like to continue using or maintaining MGT, you are welcome to fork the repository: ' +
        'https://github.com/microsoftgraph/microsoft-graph-toolkit'
    );
  } catch {
    // Never allow logging to break a host application.
  }
};
