/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement as litElement } from 'lit/decorators.js';
import { customElementHelper } from '../components/customElementHelper';

/**
 * This is a wrapper decorator for `customElement` from `lit`
 * It adds the appropriate prefix to the provided tagName calls the wrapped decorator
 * This decorator should be used in place of the unwrapped version from lit in all cases.
 *
 * @param tagName the base name for the custom element tag
 */
export const customElement = (tagName: string): ((classOrDescriptor: unknown) => any) => {
  const mgtTagName = `${customElementHelper.prefix}-${tagName}`;
  const mgtElement = customElements.get(mgtTagName);
  const unknownVersion = 'Unknown possibly <3.0.0';
  const version = (mgtElement as any).version || unknownVersion;
  if (mgtElement) {
    // tslint:disable-next-line: no-console
    console.error(`Element with tag name ${mgtTagName} already exists, using for using ${mgtElement.name}@${version}`);
    return (classOrDescriptor: unknown) => {
      // tslint:disable-next-line: no-console
      console.error(`${mgtTagName} for wants to use v${(classOrDescriptor as any).version || unknownVersion}`);
      return classOrDescriptor;
    };
  }
  return litElement(mgtTagName);
};
