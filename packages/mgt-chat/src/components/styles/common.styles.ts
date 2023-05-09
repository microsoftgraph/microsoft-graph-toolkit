/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { mergeStyleSets, mergeStyles } from '@fluentui/react/lib/Styling';

const classNames = {
  buttonIcon: 'button-icon',
  icon: 'icon',
  iconFilled: 'icon-filled',
  iconUnfilled: 'icon-unfilled'
};

const iconOnlyButtonStyle = mergeStyles({
  width: '18px',
  height: '18px'
});

const buttonIconStyle = {};

const colors = {
  iconRest: '#616161',
  iconHover: '#5b5fc7'
};

const buttonIconStyles = mergeStyleSets({
  button: [classNames.buttonIcon, buttonIconStyle],

  icon: [
    classNames.icon,
    {
      fill: colors.iconRest,
      [`.${classNames.buttonIcon}:hover &`]: {
        fill: colors.iconHover,
        color: colors.iconHover
      }
    }
  ],

  iconFilled: [
    classNames.iconFilled,
    {
      display: 'none',
      [`.${classNames.buttonIcon}:hover &`]: {
        display: 'block'
      }
    }
  ],

  iconUnfilled: [
    classNames.iconUnfilled,
    {
      display: 'block',
      [`.${classNames.buttonIcon}:hover &`]: {
        display: 'none'
      }
    }
  ]
});

export { buttonIconStyles, buttonIconStyle, iconOnlyButtonStyle };
