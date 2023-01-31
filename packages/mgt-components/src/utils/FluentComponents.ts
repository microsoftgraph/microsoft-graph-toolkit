/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { provideFluentDesignSystem } from '@fluentui/web-components';

const designSystem = provideFluentDesignSystem();

export const registerFluentComponents = (...fluentComponents) => {
  if (!fluentComponents || !fluentComponents.length) {
    return;
  }

  for (const component of fluentComponents) {
    designSystem.register(component());
  }
};
