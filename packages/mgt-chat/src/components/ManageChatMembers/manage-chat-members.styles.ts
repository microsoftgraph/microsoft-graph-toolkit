/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { mergeStyles } from '@fluentui/react';
const styles = {
  actionIcon: mergeStyles({
    height: '18px',
    width: '18px'
  }),
  addMembers: mergeStyles({
    display: 'flex',
    flexDirection: 'column'
  }),
  buttonRow: mergeStyles({
    display: 'flex',
    flexDirection: 'row'
  }),
  container: mergeStyles({
    display: 'flex',
    flexDirection: 'column',
    width: 'max-content',
    height: 'max-content',
    padding: '4px 12px 12px 12px'
  }),
  historyInput: mergeStyles({ width: '48px' }),
  listItem: mergeStyles({
    listStyleType: 'none',
    width: '100%'
  }),
  manageMemberButton: mergeStyles({
    color: '#616161',
    width: 'min-content'
  }),
  memberItem: mergeStyles({
    paddingTop: '8px'
  }),
  memberList: mergeStyles({
    fontWeight: 800,
    gridGap: '8px',
    marginBlock: '0px',
    paddingInline: '0px'
  }),
  fullWidth: mergeStyles({
    width: '100%'
  }),
  popover: mergeStyles({
    paddingLeft: '0',
    paddingRight: '0',
    paddingTop: '0',
    paddingBottom: '0'
  }),
  option: mergeStyles({
    display: 'flex',
    alignItems: 'center',
    columnGap: '4px'
  }),
  triggerButton: mergeStyles({
    minWidth: 'max-content',
    width: 'max-content'
  })
};
export { styles };
