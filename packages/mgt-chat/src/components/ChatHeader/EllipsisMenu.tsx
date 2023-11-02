/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
import {
  Button,
  Menu,
  MenuItemLink,
  MenuList,
  MenuPopover,
  MenuProps,
  MenuTrigger,
  makeStyles
} from '@fluentui/react-components';
import { MoreHorizontal24Filled, MoreHorizontal24Regular, bundleIcon } from '@fluentui/react-icons';
import { createSvgIcon } from '@fluentui/react-northstar';
import React from 'react';

const TeamsIcon = createSvgIcon({
  svg: () => (
    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 20 20" fill="none">
      <path
        d="M17.0797 8H14.1797L12.4697 8.82V12.89C12.5119 13.5868 12.8184 14.2412 13.3267 14.7197C13.8349 15.1981 14.5067 15.4645 15.2047 15.4645C15.9028 15.4645 16.5745 15.1981 17.0828 14.7197C17.5911 14.2412 17.8976 13.5868 17.9397 12.89V8.82C17.9399 8.70887 17.9174 8.59886 17.8737 8.49668C17.83 8.39449 17.766 8.30225 17.6856 8.22556C17.6052 8.14887 17.51 8.08933 17.4058 8.05055C17.3017 8.01178 17.1907 7.99458 17.0797 8Z"
        fill="#115EA3"
      />
      <path
        d="M16 7C17.1046 7 18 6.10457 18 5C18 3.89543 17.1046 3 16 3C14.8954 3 14 3.89543 14 5C14 6.10457 14.8954 7 16 7Z"
        fill="#115EA3"
      />
      <path
        d="M6.82 8H14.18C14.3975 8 14.606 8.08639 14.7598 8.24017C14.9136 8.39395 15 8.60252 15 8.82V13.5C15 14.6935 14.5259 15.8381 13.682 16.682C12.8381 17.5259 11.6935 18 10.5 18C9.30653 18 8.16193 17.5259 7.31802 16.682C6.47411 15.8381 6 14.6935 6 13.5V8.82C6 8.60252 6.08639 8.39395 6.24017 8.24017C6.39395 8.08639 6.60252 8 6.82 8Z"
        fill="#B4D6FA"
      />
      <path
        d="M11 8C12.6569 8 14 6.65685 14 5C14 3.34315 12.6569 2 11 2C9.34315 2 8 3.34315 8 5C8 6.65685 9.34315 8 11 8Z"
        fill="#B4D6FA"
      />
      <path
        opacity="0.5"
        d="M6.00003 8.82V13.5C5.99696 14.3892 6.25799 15.2593 6.75003 16H10.97C11.5005 16 12.0092 15.7893 12.3842 15.4142C12.7593 15.0391 12.97 14.5304 12.97 14V8H6.75003C6.5449 8.01757 6.35389 8.11166 6.21494 8.26358C6.07599 8.4155 5.99928 8.61412 6.00003 8.82Z"
        fill="#242424"
      />
      <path
        opacity="0.5"
        d="M11.9699 6H8.16992C8.3332 6.46135 8.60686 6.87575 8.96704 7.20706C9.32722 7.53837 9.76299 7.77654 10.2363 7.9008C10.7097 8.02505 11.2062 8.03162 11.6827 7.91993C12.1592 7.80823 12.6011 7.58167 12.9699 7.26V7C12.9699 6.73478 12.8646 6.48043 12.677 6.29289C12.4895 6.10536 12.2351 6 11.9699 6Z"
        fill="#242424"
      />
      <path
        opacity="0.1"
        d="M6.00003 8.82V13.5C5.99696 14.3892 6.25799 15.2593 6.75003 16H10.97C11.5005 16 12.0092 15.7893 12.3842 15.4142C12.7593 15.0391 12.97 14.5304 12.97 14V8H6.75003C6.5449 8.01757 6.35389 8.11166 6.21494 8.26358C6.07599 8.4155 5.99928 8.61412 6.00003 8.82Z"
        fill="white"
      />
      <path
        opacity="0.1"
        d="M11.9699 6H8.16992C8.3332 6.46135 8.60686 6.87575 8.96704 7.20706C9.32722 7.53837 9.76299 7.77654 10.2363 7.9008C10.7097 8.02505 11.2062 8.03162 11.6827 7.91993C12.1592 7.80823 12.6011 7.58167 12.9699 7.26V7C12.9699 6.73478 12.8646 6.48043 12.677 6.29289C12.4895 6.10536 12.2351 6 11.9699 6Z"
        fill="white"
      />
      <path
        d="M11 5H3C2.44772 5 2 5.44772 2 6V14C2 14.5523 2.44772 15 3 15H11C11.5523 15 12 14.5523 12 14V6C12 5.44772 11.5523 5 11 5Z"
        fill="#115EA3"
      />
      <path d="M10 8H8V13H7V8H5V7H10V8Z" fill="white" />
    </svg>
  ),
  displayName: 'TeamsIcon'
});

const TeamsEllipsisIcon = bundleIcon(MoreHorizontal24Filled, MoreHorizontal24Regular);

const menuProps: Partial<MenuProps> = {
  inline: true,
  hasIcons: true,
  closeOnScroll: true,
  positioning: 'below-start'
};

const ellipsisMenuStyles = makeStyles({
  menuPopover: {
    zIndex: 2,

    '& .fui-MenuItemLink__content': {
      fontFamily: 'Segoe UI',
      fontSize: '12px',
      fontStyle: 'normal',
      fontWeight: '600'
    }
  }
});
interface MenuItemsProps {
  chatWebUrl?: string;
}

const EllipsisMenu = (props: MenuItemsProps) => {
  const styles = ellipsisMenuStyles();

  return (
    <Menu {...menuProps}>
      <MenuTrigger disableButtonEnhancement>
        <Button appearance="transparent" icon={<TeamsEllipsisIcon />} />
      </MenuTrigger>

      <MenuPopover className={styles.menuPopover}>
        <MenuList>
          <MenuItemLink icon={<TeamsIcon />} target="_blank" href={props.chatWebUrl!}>
            Go to Microsoft Teams
          </MenuItemLink>
        </MenuList>
      </MenuPopover>
    </Menu>
  );
};

export { EllipsisMenu };
