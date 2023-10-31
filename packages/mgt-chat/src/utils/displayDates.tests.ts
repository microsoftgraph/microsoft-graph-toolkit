/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { onDisplayDateTimeString } from './displayDates';

describe('displayDates - onDisplayDateTimeString', () => {
  it('shoulder render date object as string with am', () => {
    const amDate = new Date('Tue Oct 31 2023 07:08:40 GMT+0300 (East Africa Time)');
    const dateString = onDisplayDateTimeString(amDate);
    expect(dateString).toEqual('2023-10-31 4:08 a.m.');
  });

  it('shoulder render date object as string with pm', () => {
    const pmDate = new Date('Tue Oct 31 2023 17:08:40 GMT+0300 (East Africa Time)');
    const dateString = onDisplayDateTimeString(pmDate);
    expect(dateString).toEqual('2023-10-31 2:08 p.m.');
  });
});
