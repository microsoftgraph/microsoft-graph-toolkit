/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElementHelper } from '../components/customElementHelper';

export const buildComponentName = (tagBase: string) => `${customElementHelper.prefix}-${tagBase}`;

export const registerComponent = (
  tagBase: string,
  constructor: CustomElementConstructor,
  options?: ElementDefinitionOptions
) => {
  const tagName = buildComponentName(tagBase);
  if (!customElements.get(tagName)) customElements.define(tagName, constructor, options);
};
