/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { onDisplayDateTimeString } from './displayDates';
import { expect } from '@open-wc/testing';

describe('displayDates - onDisplayDateTimeString', () => {
  it('shoulder render date object as string with am', async () => {
    const amDate = new Date('Tue Oct 31 2023 07:08:40 GMT+0300 (East Africa Time)');
    const dateString = onDisplayDateTimeString(amDate);
    await expect(dateString).equal('2023-10-31 7:08 a.m.');
  });

  it('shoulder render date object as string with pm', async () => {
    const pmDate = new Date('Tue Oct 31 2023 17:08:40 GMT+0300 (East Africa Time)');
    const dateString = onDisplayDateTimeString(pmDate);
    await expect(dateString).equal('2023-10-31 5:08 p.m.');
  });
});
