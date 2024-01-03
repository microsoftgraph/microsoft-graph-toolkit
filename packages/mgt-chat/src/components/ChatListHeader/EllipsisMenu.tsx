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
import React from 'react';
import { MenuItem } from './MenuItem';

const EllipsisIcon = bundleIcon(MoreHorizontal24Filled, MoreHorizontal24Regular);

const menuProps: Partial<MenuProps> = {
  inline: true,
  hasIcons: true,
  closeOnScroll: true,
  positioning: 'below-start'
};

const ellipsisMenuStyles = makeStyles({
  menuPopover: {
    zIndex: 10,
    '& .fui-MenuItemLink__content': {
      fontFamily: 'Segoe UI',
      fontSize: '12px',
      fontStyle: 'normal',
      fontWeight: '600'
    }
  }
});
export interface IChatListMenuItemsProps {
  menuItems?: MenuItem[];
}

const EllipsisMenu = (props: IChatListMenuItemsProps) => {
  const styles = ellipsisMenuStyles();

  const menuItems: MenuItem[] = props.menuItems === undefined ? [] : props.menuItems;

  return (
    <Menu {...menuProps}>
      <MenuTrigger disableButtonEnhancement>
        <Button appearance="transparent" icon={<EllipsisIcon />} />
      </MenuTrigger>
      <MenuPopover className={styles.menuPopover}>
        <MenuList>
          {menuItems.map(menuItem => (
            <MenuItemLink href="" onSelect={menuItem.onSelected}>
              {menuItem.displayText}
            </MenuItemLink>
          ))}
        </MenuList>
      </MenuPopover>
    </Menu>
  );
};

export { EllipsisMenu };
