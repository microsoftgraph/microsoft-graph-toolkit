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
  SkypeArrow,

  SmallEmail,

  SmallChat,

  ExpandDown,

  Overview,

  Send
}

import { html } from 'lit-element';

/**
 * returns an svg
 * @param svgIcon defined by name
 * @param color hex value
 */
export function getSvg(svgIcon: SvgIcon, color?: string) {
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

    case SvgIcon.SmallEmail:
      return html`
        <svg width="14" height="10" viewBox="0 0 14 10" xmlns="http://www.w3.org/2000/svg">
          <path
            fill-rule="evenodd"
            clip-rule="evenodd"
            d="M11.8473 1H2.04886L6.47364 4.18916C6.64273 4.31103 6.86969 4.31522 7.04316 4.19968L11.8473 1ZM1 1.47671V9H13V1.43376L7.59749 5.03198C7.07706 5.3786 6.39621 5.36601 5.88894 5.0004L1 1.47671ZM0 1C0 0.447715 0.447715 0 1 0H13C13.5523 0 14 0.447715 14 1V9C14 9.55228 13.5523 10 13 10H1C0.447716 10 0 9.55229 0 9V1Z"
          />
        </svg>
      `;

    case SvgIcon.SmallChat:
      return html`
        <svg width="13" height="13" viewBox="0 0 13 13" xmlns="http://www.w3.org/2000/svg">
          <path
            fill-rule="evenodd"
            clip-rule="evenodd"
            d="M9.5781 8.26403L9.57811 8.26403C9.68611 8.29867 9.8455 8.36364 10.0125 8.50479C10.1405 8.61294 10.2409 8.73883 10.3235 8.86465L11.9634 10.9944C11.9701 10.9924 11.9753 10.9904 11.9785 10.9889C11.9841 10.9864 11.9918 10.9823 12 10.9768V1.32284C12 1.18078 11.8731 1 11.6207 1H1.37926C1.12692 1 1 1.18078 1 1.32284V7.37377C1 7.45357 1.01415 7.49036 1.02102 7.50507C1.02778 7.51955 1.04342 7.54689 1.09159 7.58705L1.10485 7.5981L1.11771 7.6096C1.13526 7.62529 1.21707 7.69076 1.33937 7.76027C1.46122 7.82952 1.58119 7.87864 1.67944 7.89966L1.69102 7.90214L1.691 7.9022C3.32106 8.27116 6.2626 8.27688 8.67896 8.18036L8.69908 8.17955L8.71921 8.17956L8.7627 8.17954C9.01362 8.17932 9.31313 8.17907 9.5781 8.26403ZM11.2376 11.6908L9.50506 9.44081C9.39493 9.26422 9.32445 9.23285 9.27276 9.21627C9.17534 9.18504 9.0401 9.17966 8.71888 9.17956C6.31879 9.27543 3.24831 9.27999 1.47024 8.87753C1.01314 8.77974 0.600449 8.48852 0.451238 8.35513C0.15593 8.10893 0 7.78958 0 7.37377V1.32284C0 0.593147 0.610664 0 1.37926 0H11.6207C12.3893 0 13 0.593147 13 1.32284V11.0441C13 11.4686 12.6828 11.7689 12.3881 11.9012C12.1048 12.0284 11.673 12.0728 11.3387 11.7993C11.3003 11.7679 11.2678 11.7301 11.2376 11.6908ZM3 3.5C3 3.22386 3.22386 3 3.5 3H9.5C9.77614 3 10 3.22386 10 3.5C10 3.77614 9.77614 4 9.5 4H3.5C3.22386 4 3 3.77614 3 3.5ZM3.5 5C3.22386 5 3 5.22386 3 5.5C3 5.77614 3.22386 6 3.5 6H6.5C6.77614 6 7 5.77614 7 5.5C7 5.22386 6.77614 5 6.5 5H3.5Z"
          />
        </svg>
      `;

    case SvgIcon.ExpandDown: {
      return html`
        <svg width="15" height="8" viewBox="0 0 15 8" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M14 1L7.5 7L1 1" stroke="#3078CD" />
        </svg>
      `;
    }

    case SvgIcon.Overview:
      return html`
        <svg xmlns="http://www.w3.org/2000/svg">
          <path
            d="M12 4V9H2V4H12ZM11 8V5H3V8H11ZM13 4H18V9H13V4ZM17 8V5H14V8H17ZM8 15V10H18V15H8ZM9 11V14H17V11H9ZM2 15V10H7V15H2ZM3 11V14H6V11H3Z"
          />
        </svg>
      `;

    case SvgIcon.Send:
      return html`
        <svg xmlns="http://www.w3.org/2000/svg">
          <path
            d="M4.27144 8.99999L1.72572 2.45387C1.54854 1.99826 1.9928 1.56256 2.43227 1.71743L2.50153 1.74688L16.0015 8.49688C16.3902 8.69122 16.4145 9.22336 16.0744 9.45992L16.0015 9.50311L2.50153 16.2531C2.0643 16.4717 1.58932 16.0697 1.70282 15.6178L1.72572 15.5461L4.27144 8.99999L1.72572 2.45387L4.27144 8.99999ZM3.3028 3.4053L5.25954 8.43705L10.2302 8.43749C10.515 8.43749 10.7503 8.64911 10.7876 8.92367L10.7927 8.99999C10.7927 9.28476 10.5811 9.52011 10.3065 9.55736L10.2302 9.56249L5.25954 9.56205L3.3028 14.5947L14.4922 8.99999L3.3028 3.4053Z"
          />
        </svg>
      `;
  }
}
