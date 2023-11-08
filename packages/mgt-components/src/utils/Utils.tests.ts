/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { getRelativeDisplayDate } from './Utils';
import { expect } from '@open-wc/testing';

describe('Utils - getRelativeDisplayDate', () => {
  it('should render the date today in AM format', async () => {
    const today = new Date();
    today.setHours(10, 10);
    const result = getRelativeDisplayDate(today);
    await expect(result).equal('10:10 AM');
  });

  it('should render the date today in PM format', async () => {
    const today = new Date();
    today.setHours(15, 10);
    const result = getRelativeDisplayDate(today);
    await expect(result).equal('3:10 PM');
  });

  it('should show the date in more than two weeks ago format', async () => {
    const today = new Date();
    const twoWeeksAgo = today.getDate() - 15;
    today.setMonth(1);
    today.setDate(twoWeeksAgo);
    const result = getRelativeDisplayDate(today);
    await expect(result).equal('1/24/2023');
  });
});
