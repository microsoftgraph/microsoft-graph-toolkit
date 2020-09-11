/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * Defines icon used by svgHelper
 *
 * @export
 * @enum {number}
 */
export enum SvgIcon {
  /**
   * Arrow Icon pointing right
   */
  ArrowRight,

  /**
   * Arrow Icon pointing down
   */
  ArrowDown,

  /**
   * Icon separates team from channel in selection
   */
  TeamSeparator,

  /**
   * Search icon
   */
  Search,

  /**
   * Skype Arrow icon (out of office status)
   */
  SkypeArrow
}

import { html } from 'lit-element';

/**
 * returns an svg
 * @param svgIcon defined by name
 * @param color hex value
 */
export function getSvg(svgIcon: SvgIcon, color: string) {
  switch (svgIcon) {
    case SvgIcon.ArrowRight:
      return html`
        <svg width="12" height="12" viewBox="0 0 12 12" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M8 7L4.46481 10.5359L4.46481 7L4.46481 3.46413L8 7Z" />
        </svg>
      `;

    case SvgIcon.ArrowDown:
      return html`
        <svg width="12" height="12" viewBox="0 0 12 12" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M6 9L2.46447 5.46447H6H9.53553L6 9Z" />
        </svg>
      `;

    case SvgIcon.TeamSeparator:
      return html`
        <svg width="6" height="10" viewBox="0 0 6 10" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path
            fill-rule="evenodd"
            clip-rule="evenodd"
            d="M5.70711 5L1.49999 9.20711L0.792886 8.50001L4.29289 5L0.792887 1.49999L1.49999 0.792885L5.70711 5Z"
            fill=${color}
          />
        </svg>
      `;

    case SvgIcon.Search:
      return html`
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" xmlns="http://www.w3.org/2000/svg">
          <g clip-path="url(#clip0)">
            <circle cx="5.36377" cy="5.36396" r="4" transform="rotate(45 5.36377 5.36396)" stroke="#B3B0AD" />
            <path
              d="M8.19189 7.48529L12.7881 12.0815C12.9834 12.2767 12.9834 12.5933 12.7881 12.7886V12.7886C12.5928 12.9839 12.2762 12.9839 12.081 12.7886L7.48479 8.1924L8.19189 7.48529Z"
              fill="#B3B0AD"
            />
          </g>
          <defs>
            <clipPath id="clip0">
              <rect width="14" height="14" fill="white" />
            </clipPath>
          </defs>
        </svg>
      `;

    case SvgIcon.SkypeArrow:
      return html`
        <svg viewBox="0 0 12 10" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path
            d="M3.95184 0.480534C4.23385 0.10452 4.70926 -0.0724722 5.1685 0.0275755C5.62775 0.127623 5.98645 0.486329 6.0865 0.945575C6.18655 1.40482 6.00956 1.88023 5.63354 2.16224L4.07196 3.72623H10.7988C11.4622 3.72623 12 4.26403 12 4.92744C12 5.59086 11.4622 6.12866 10.7988 6.12866H4.07196L5.63114 7.68784C6.0955 8.15225 6.0955 8.90515 5.63114 9.36955C5.51655 9.48381 5.38119 9.57514 5.23234 9.63862C5.09341 9.69857 4.94399 9.73042 4.79269 9.73232C4.63498 9.73233 4.4789 9.70046 4.33382 9.63862C4.18765 9.57669 4.05593 9.48507 3.94703 9.36955L0.343377 5.7659C-0.114459 5.29881 -0.114459 4.55128 0.343377 4.08419L3.95184 0.480534Z"
            fill="#B4009E"
          />
        </svg>
      `;
  }
}
