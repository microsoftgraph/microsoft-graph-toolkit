/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { fixture, html, expect } from '@open-wc/testing';
import './mgt-spinner';
import { registerMgtSpinnerComponent } from './mgt-spinner';

let spinner: Element;
describe('mgt-spinner tests', () => {
  // before(() => reg);
  beforeEach(async () => {
    registerMgtSpinnerComponent();
    spinner = await fixture(html`<mgt-spinner></mgt-spinner>`);
  });

  it('should render', async () => {
    await expect(spinner).shadowDom.to.equal('<fluent-progress-ring title="spinner"></fluent-progress-ring>');
  });
});
