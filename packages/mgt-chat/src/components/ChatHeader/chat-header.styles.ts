/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { mergeStyles } from '@fluentui/react';
const styles = {
  chatHeader: mergeStyles({
    display: 'flex',
    alignItems: 'center',
    columnGap: '4px',
    lineHeight: '32px',
    fontSize: '18px',
    fontWeight: 700
  })
};
export { styles };
