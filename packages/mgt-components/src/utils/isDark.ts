/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { SwatchRGB, baseLayerLuminance, isDark } from '@fluentui/web-components';

/**
 * Utility to help quickly determine if an element is dark based fluentui theme
 *
 * @param element HTMLElement to check if dark
 * @returns true if the element is dark
 */
export const isElementDark = (element: HTMLElement) => {
  const luminance = baseLayerLuminance.getValueFor(element);
  return isDark(SwatchRGB.create(luminance, luminance, luminance));
};
