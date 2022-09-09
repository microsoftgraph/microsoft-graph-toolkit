/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement as litElement } from 'lit/decorators';
import { customElementHelper } from '../components/customElementHelper';

/**
 * This is a wrapper decorator for `customElement` from `lit`
 * It adds the appropriate prefix to the provided tagName calls the wrapped decorator
 * This decorator should be used in place of the unwrapped version from lit in all cases.
 *
 * @param tagName the base name for the custom element tag
 */
export const customElement = (tagName: string): ((classOrDescriptor: unknown) => any) => {
  return litElement(`${customElementHelper.prefix}-${tagName}`);
};
