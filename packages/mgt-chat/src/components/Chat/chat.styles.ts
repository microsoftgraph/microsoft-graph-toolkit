/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { mergeStyleSets } from '@fluentui/react';
const chatStyles = mergeStyleSets({
  chat: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    overflow: 'auto'
  },

  chatMessages: {
    height: 'auto',
    overflow: 'auto'
  },

  chatInput: {
    overflow: 'unset'
  },
  fullHeight: {
    height: '100%'
  },
  spinnerContainer: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    height: 'calc(100vh - 300px)'
  }
});

export { chatStyles };
