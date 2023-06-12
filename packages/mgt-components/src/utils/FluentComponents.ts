/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { provideFluentDesignSystem } from '@fluentui/web-components';

/**
 * Provides a design system to the fluent components
 */
const designSystem = provideFluentDesignSystem();

/**
 * Registers fluent components to the design system
 *
 * @param fluentComponents array of fluent components to register
 * @returns
 */
export const registerFluentComponents = (...fluentComponents: (() => unknown)[]) => {
  if (!fluentComponents?.length) {
    return;
  }

  for (const component of fluentComponents) {
    designSystem.register(component());
  }
};
