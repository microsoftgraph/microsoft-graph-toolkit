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
  items: MenuItem[];
}

const EllipsisMenu = (props: MenuItemsProps) => {
  const styles = ellipsisMenuStyles();

  const menuItems: MenuItem[] = props.items;

  return (
    <Menu {...menuProps}>
      <MenuTrigger disableButtonEnhancement>
        <Button appearance="transparent" icon={<EllipsisIcon />} />
      </MenuTrigger>
      <MenuPopover className={styles.menuPopover}>
        <MenuList>
          {menuItems.map(menuItem => (
            <MenuItemLink href="" icon={menuItem.icon}>
              {menuItem.displayText}
            </MenuItemLink>
          ))}
        </MenuList>
      </MenuPopover>
    </Menu>
  );
};

export { EllipsisMenu };
