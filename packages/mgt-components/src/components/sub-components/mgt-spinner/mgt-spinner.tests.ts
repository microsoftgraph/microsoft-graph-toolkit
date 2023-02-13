/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { screen } from 'testing-library__dom';
import { fixture } from '@open-wc/testing-helpers';
import './mgt-spinner';

let spinner: Element;
describe('mgt-spinner tests', () => {
  beforeEach(async () => {
    spinner = await fixture('<mgt-spinner></mgt-spinner>');
  });

  it('should render', async () => {
    const spinnerHtml = await screen.findByTitle('spinner');
    expect(spinnerHtml).not.toBeNull();
    expect(spinner.shadowRoot.innerHTML).toContain('<fluent-progress-ring title="spinner"></fluent-progress-ring>');
  });
});
