import {
  Button,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuProps,
  MenuTrigger,
  makeStyles
} from '@fluentui/react-components';
import { MoreHorizontal24Filled, MoreHorizontal24Regular, bundleIcon } from '@fluentui/react-icons';
import React from 'react';
import { ChatListMenuItem } from './ChatListMenuItem';

const EllipsisIcon = bundleIcon(MoreHorizontal24Filled, MoreHorizontal24Regular);

const menuProps: Partial<MenuProps> = {
  inline: true,
  hasIcons: false,
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
  menuItems?: ChatListMenuItem[];
}

const EllipsisMenu = (props: IChatListMenuItemsProps) => {
  const styles = ellipsisMenuStyles();

  const menuItems: ChatListMenuItem[] = props.menuItems === undefined ? [] : props.menuItems;

  return (
    <Menu {...menuProps}>
      <MenuTrigger disableButtonEnhancement>
        <Button appearance="transparent" icon={<EllipsisIcon />} />
      </MenuTrigger>
      <MenuPopover className={styles.menuPopover}>
        <MenuList>
          {menuItems.map((menuItem, index) => (
            <MenuItem key={index} content={menuItem.displayText} onClick={menuItem.onClick} />
          ))}
        </MenuList>
      </MenuPopover>
    </Menu>
  );
};

export { EllipsisMenu };
