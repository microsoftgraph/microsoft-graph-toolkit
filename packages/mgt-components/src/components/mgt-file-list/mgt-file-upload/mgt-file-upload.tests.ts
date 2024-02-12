/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { fixture, html, expect } from '@open-wc/testing';
import { MockProvider, Providers } from '@microsoft/mgt-element';
import { MgtFileUpload, registerMgtFileUploadComponent } from './mgt-file-upload';

describe('mgt-file-upload - tests', () => {
  registerMgtFileUploadComponent();
  Providers.globalProvider = new MockProvider(true);

  it('should render', async () => {
    const mgtFileUpload = await fixture(html`<mgt-file-upload></mgt-file-upload>`);

    await expect(mgtFileUpload).shadowDom.to.equal(`
    <div
        class="file-upload-dialog"
        id="file-upload-dialog">
        <fluent-dialog
          class="file-upload-dialog-content"
          modal="true">
          <span
            class="file-upload-dialog-close"
            id="file-upload-dialog-close">
          </span>
          <div class="file-upload-dialog-content-text">
            <h2 class="file-upload-dialog-title"></h2>
            <div></div>
            <fluent-checkbox
              aria-checked="false"
              aria-disabled="false"
              aria-required="false"
              class="file-upload-dialog-check"
              id="file-upload-dialog-check"
              role="checkbox"
              tabindex="0">
            </fluent-checkbox>
          </div>
          <div class="file-upload-dialog-editor">
            <fluent-button
              appearance="accent"
              class="accent file-upload-dialog-ok">
            </fluent-button>
            <fluent-button
              appearance="outline"
              class="file-upload-dialog-cancel outline">
            </fluent-button>
          </div>
        </fluent-dialog>
      </div>
      <div id="file-upload-border">
      </div>
      <div class="file-upload-area-button">
        <input
          aria-label="file upload input"
          id="file-upload-input"
          multiple=""
          tabindex="-1"
          title="File upload button"
          type="file">
        <fluent-button
          appearance="accent"
          class="accent file-upload-button"
          label="File upload button">
          <span slot="start"></span>
          <span class="upload-text">
            Upload Files
          </span>
        </fluent-button>
      </div>
      <div class="file-upload-template"></div>
    `);
  });

  it('has required scopes', () => {
    expect(MgtFileUpload.requiredScopes).to.have.members([
      'files.readwrite',
      'files.readwrite.all',
      'sites.readwrite.all'
    ]);
  });
});
