import { SwatchRGB, baseLayerLuminance, isDark } from '@fluentui/web-components';

export const isElementDark = (element: HTMLElement) => {
  const luminance = baseLayerLuminance.getValueFor(element);
  return isDark(SwatchRGB.create(luminance, luminance, luminance));
};
