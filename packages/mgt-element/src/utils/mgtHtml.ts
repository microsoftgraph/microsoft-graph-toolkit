import { HTMLTemplateResult, html } from 'lit';
import { customElementHelper } from '../components/customElementHelper';

/**
 * stringsCache ensures that the same TemplateStringsArray object is returned on subsequent calls
 * This is needed as lit-html internally uses the TemplateStringsArray object as a cache key.
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
const rewriteStrings = (strings: ReadonlyArray<string>, matcher: RegExp): ReadonlyArray<string> => {
  const temp: string[] = [];
  const newPrefix = `${customElementHelper.prefix}-`;
  for (const s of strings) {
    temp.push(s.replace(matcher, '$1' + newPrefix));
  }
  return temp;
};

/**
 * Generates a template literal tag function that returns an HTMLTemplateResult.
 */
const tag = (strings: TemplateStringsArray, ...values: unknown[]): HTMLTemplateResult => {
  // re-write <mgt-([a-z]+) if necessary
  const prefix = customElementHelper.prefix;
  if (prefix !== customElementHelper.defaultPrefix) {
    let cached = stringsCache.get(strings);
    if (!Boolean(cached)) {
      const matcher = new RegExp('(</?)mgt-(?!contoso-)');
      cached = Object.assign(rewriteStrings(strings, matcher), { raw: rewriteStrings(strings.raw, matcher) });
      stringsCache.set(strings, cached);
    }
    strings = cached;
  }

  return html(strings, ...values);
};

/**
 * Interprets a template literal and dynamically rewrites `<mtg-` tags with the
 * configured disambiguation if necessary.
 *
 * ```ts
 * const header = (title: string) => mgtHtml`<mgt-flyout>${title}</mgt-flyout>`;
 * ```
 *
 * The `mgtHtml` tag is a wrapper for the `html` tag from `lit` which provides for dynamic tag re-writing
 */
export const mgtHtml = tag;
