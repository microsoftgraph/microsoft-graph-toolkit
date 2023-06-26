/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { HTMLTemplateResult, html } from 'lit';
import { customElementHelper } from '../components/customElementHelper';

/**
 * stringsCache ensures that the same TemplateStringsArray object is returned on subsequent calls
 * This is needed as lit-html internally uses the TemplateStringsArray object as a cache key.
 *
 * @type {WeakMap}
 */
const stringsCache = new WeakMap<TemplateStringsArray, TemplateStringsArray>();

/**
 * Rewrites strings in an array using the supplied matcher RegExp
 * Assumes that the RegExp returns a group on matches
 *
 * @param strings The array of strings to be re-written
 * @param matcher A RegExp to be used for matching strings for replacement
 */
const rewriteStrings = (strings: readonly string[], matcher: RegExp, replacement: string): readonly string[] => {
  const temp: string[] = [];
  for (const s of strings) {
    temp.push(s.replace(matcher, replacement));
  }
  return temp;
};

/**
 * Generates a template literal tag function that returns an HTMLTemplateResult.
 */
const tag = (strings: TemplateStringsArray, ...values: unknown[]): HTMLTemplateResult => {
  // re-write <mgt-([a-z]+) if necessary
  if (customElementHelper.isDisambiguated) {
    let cached = stringsCache.get(strings);
    if (!cached) {
      const matcher = new RegExp('(</?)mgt-(?!' + customElementHelper.disambiguation + '-)');
      const newPrefix = `$1${customElementHelper.prefix}-`;
      cached = Object.assign(rewriteStrings(strings, matcher, newPrefix), {
        raw: rewriteStrings(strings.raw, matcher, newPrefix)
      });
      stringsCache.set(strings, cached);
    }
    strings = cached;
  }

  return html(strings, ...values);
};

/**
 * Interprets a template literal and dynamically rewrites `<mgt-` tags with the
 * configured disambiguation if necessary.
 *
 * ```ts
 * const header = (title: string) => mgtHtml`<mgt-flyout>${title}</mgt-flyout>`;
 * ```
 *
 * The `mgtHtml` tag is a wrapper for the `html` tag from `lit` which provides for dynamic tag re-writing
 */
export const mgtHtml = tag;
