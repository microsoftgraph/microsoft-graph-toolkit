/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/* eslint-disable @typescript-eslint/no-unused-expressions */

import { expect } from '@open-wc/testing';
import { sanitizeSummary } from './Utils';

describe('sanitizeSummary', () => {
  it('should return falsy values unchanged', async () => {
    expect(sanitizeSummary(null)).to.be.null;
    expect(sanitizeSummary(undefined)).to.be.undefined;
    await expect(sanitizeSummary('')).to.equal('');
  });

  it('should convert <ddd/> to ellipsis', async () => {
    await expect(sanitizeSummary('hello<ddd/>world')).to.equal('hello...world');
  });

  it('should convert <c0> and </c0> to <b> and </b>', async () => {
    await expect(sanitizeSummary('a <c0>match</c0> b')).to.equal('a <b>match</b> b');
  });

  it('should handle all proprietary tags together', async () => {
    const input = 'Result <c0>keyword</c0> in document<ddd/>';
    const expected = 'Result <b>keyword</b> in document...';
    await expect(sanitizeSummary(input)).to.equal(expected);
  });

  it('should strip <script> tags', () => {
    const input = '<script>alert(1)</script>';
    expect(sanitizeSummary(input)).to.not.contain('<script');
    expect(sanitizeSummary(input)).to.not.contain('alert');
  });

  it('should strip <img> with onerror handler', () => {
    const input = '<img src=x onerror="alert(document.domain)">';
    expect(sanitizeSummary(input)).to.not.contain('<img');
    expect(sanitizeSummary(input)).to.not.contain('onerror');
  });

  it('should strip <svg> with onload handler', () => {
    const input = '<svg onload="alert(1)"><rect/></svg>';
    expect(sanitizeSummary(input)).to.not.contain('<svg');
    expect(sanitizeSummary(input)).to.not.contain('onload');
  });

  it('should strip javascript: URIs', () => {
    const input = '<a href="javascript:alert(1)">click</a>';
    expect(sanitizeSummary(input)).to.not.contain('javascript:');
    expect(sanitizeSummary(input)).to.not.contain('<a');
  });

  it('should strip event handler attributes from allowed tags', () => {
    const input = '<b onclick="alert(1)">text</b>';
    const result = sanitizeSummary(input);
    expect(result).to.contain('<b>');
    expect(result).to.not.contain('onclick');
  });

  it('should preserve allowed tags (b, em, strong, span)', async () => {
    const input = '<b>bold</b> <em>italic</em> <strong>strong</strong> <span>span</span>';
    await expect(sanitizeSummary(input)).to.equal(input);
  });

  it('should handle mixed proprietary tags and XSS payloads', () => {
    const input = 'Result <c0>match</c0> <ddd/> <img src=x onerror="alert(document.domain)"> end';
    const result = sanitizeSummary(input);
    expect(result).to.contain('<b>match</b>');
    expect(result).to.contain('...');
    expect(result).to.not.contain('<img');
    expect(result).to.not.contain('onerror');
  });
});
