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
import React, { useCallback } from 'react';
import { ChatListMenuItem } from './ChatListMenuItem';
import IChatListActions from './IChatListActions';

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

const EllipsisMenu = (
  props: IChatListMenuItemsProps & {
    actions: IChatListActions;
  }
) => {
  const styles = ellipsisMenuStyles();

  const menuItems: ChatListMenuItem[] = props.menuItems === undefined ? [] : props.menuItems;

  const clickMenuItem = useCallback((menuItem: ChatListMenuItem) => {
    menuItem.onClick(props.actions);
  }, []);

  return (
    <Menu {...menuProps}>
      <MenuTrigger disableButtonEnhancement>
        <Button appearance="transparent" icon={<EllipsisIcon />} />
      </MenuTrigger>
      <MenuPopover className={styles.menuPopover}>
        <MenuList>
          {menuItems.map((menuItem, index) => (
            <MenuItem key={index} content={menuItem.displayText} onClick={() => clickMenuItem(menuItem)} />
          ))}
        </MenuList>
      </MenuPopover>
    </Menu>
  );
};

export { EllipsisMenu };
