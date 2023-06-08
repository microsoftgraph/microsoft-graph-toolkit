/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  accentBaseColor,
  accentFillActive,
  accentFillFocus,
  accentFillHover,
  accentFillRest,
  accentForegroundActive,
  accentForegroundFocus,
  accentForegroundHover,
  accentForegroundRest,
  accentStrokeControlActive,
  accentStrokeControlFocus,
  accentStrokeControlHover,
  accentStrokeControlRest,
  baseLayerLuminance,
  foregroundOnAccentActive,
  foregroundOnAccentFocus,
  foregroundOnAccentHover,
  foregroundOnAccentRecipe,
  foregroundOnAccentRest,
  foregroundOnAccentRestLarge,
  neutralBaseColor,
  StandardLuminance,
  SwatchRGB
} from '@fluentui/web-components';
// @microsoft/fast-colors is a transitive dependency of @fluentui/web-components, no need to explicitly add it to package.json
import { parseColorHexRGB } from '@microsoft/fast-colors';

/**
 * Available predefined themes
 */
type Theme = 'light' | 'dark' | 'default' | 'contrast';

/**
 * Helper function to apply fluent ui theme to an element
 *
 * @export
 * @param {Theme} theme - theme name, if an unknown theme is provided, the light theme will be applied
 * @param {HTMLElement} [element=document.body]
 */
export const applyTheme = (theme: Theme, element: HTMLElement = document.body): void => {
  const settings = getThemeSettings(theme);
  applyColorScheme(settings, element);
};

/**
 * Simple data holder for theme settings
 */
type ColorScheme = {
  /**
   * Hex color string for accent base color
   *
   * @type {string}
   */
  accentBaseColor: string;
  /**
   * Hex color string for neutral base color
   *
   * @type {string}
   */
  neutralBaseColor: string;
  /**
   * Base layer luminance for the theme
   * in the range of 0 to 1
   *
   * @type {number}
   */
  baseLayerLuminance: number;

  /**
   * Optional function to override design tokens
   */
  designTokenOverrides?: (element: HTMLElement) => void;
};

/**
 * Helper function to apply fluent ui color scheme to an element
 *
 * @param {ColorScheme} settings
 * @param {HTMLElement} [element=document.body]
 */
const applyColorScheme = (settings: ColorScheme, element: HTMLElement = document.body): void => {
  accentBaseColor.setValueFor(element, SwatchRGB.from(parseColorHexRGB(settings.accentBaseColor)));
  neutralBaseColor.setValueFor(element, SwatchRGB.from(parseColorHexRGB(settings.neutralBaseColor)));
  baseLayerLuminance.setValueFor(element, settings.baseLayerLuminance);
  settings.designTokenOverrides?.(element);
};

/**
 * Helper function to translate theme name to theme settings
 *
 * @param {Theme} theme
 * @return {*}  {ThemeSettings}
 */
const getThemeSettings = (theme: Theme): ColorScheme => {
  switch (theme) {
    case 'contrast':
      return {
        accentBaseColor: '#7f85f5',
        neutralBaseColor: '#adadad',
        baseLayerLuminance: StandardLuminance.DarkMode
      };
    case 'default': // this is the Teams light theme
      return {
        accentBaseColor: '#5b5fc7',
        neutralBaseColor: '#616161',
        baseLayerLuminance: StandardLuminance.LightMode
      };
    case 'dark': // Both MGT default dark and Teams Dark theme
      return {
        accentBaseColor: '#479ef5',
        neutralBaseColor: '#adadad',
        baseLayerLuminance: StandardLuminance.DarkMode,
        designTokenOverrides: element => {
          accentFillRest.setValueFor(element, SwatchRGB.from(parseColorHexRGB('#115ea3')));
          accentFillHover.setValueFor(element, SwatchRGB.from(parseColorHexRGB('#0f6cbd')));
          accentFillActive.setValueFor(element, SwatchRGB.from(parseColorHexRGB('#0c3b5e')));
          accentFillFocus.setValueFor(element, SwatchRGB.from(parseColorHexRGB('#0f548c')));
          accentForegroundRest.setValueFor(element, SwatchRGB.from(parseColorHexRGB('#479EF5')));
          accentForegroundHover.setValueFor(element, SwatchRGB.from(parseColorHexRGB('#62abf5')));
          accentForegroundActive.setValueFor(element, SwatchRGB.from(parseColorHexRGB('#2886de')));
          accentForegroundFocus.setValueFor(element, SwatchRGB.from(parseColorHexRGB('#479ef5')));
          accentStrokeControlRest.setValueFor(element, SwatchRGB.from(parseColorHexRGB('#115ea3')));
          accentStrokeControlHover.setValueFor(element, SwatchRGB.from(parseColorHexRGB('#0f6cbd')));
          accentStrokeControlActive.setValueFor(element, SwatchRGB.from(parseColorHexRGB('#0c3b5e')));
          accentStrokeControlFocus.setValueFor(element, SwatchRGB.from(parseColorHexRGB('#0f548c')));
          foregroundOnAccentActive.setValueFor(element, SwatchRGB.from(parseColorHexRGB('#ffffff')));
          foregroundOnAccentRest.setValueFor(element, SwatchRGB.from(parseColorHexRGB('#ffffff')));
          foregroundOnAccentRestLarge.setValueFor(element, SwatchRGB.from(parseColorHexRGB('#ffffff')));
          foregroundOnAccentHover.setValueFor(element, SwatchRGB.from(parseColorHexRGB('#ffffff')));
          foregroundOnAccentFocus.setValueFor(element, SwatchRGB.from(parseColorHexRGB('#ffffff')));
        }
      };
    case 'light':
    default:
      return {
        accentBaseColor: '#0f6cbd',
        neutralBaseColor: '#616161',
        baseLayerLuminance: StandardLuminance.LightMode
      };
  }
};
