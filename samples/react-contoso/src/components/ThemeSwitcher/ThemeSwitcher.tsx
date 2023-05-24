import * as React from 'react';
import {
  Menu,
  MenuButton,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  webDarkTheme,
  webLightTheme
} from '@fluentui/react-components';
import { BrightnessHighRegular, WeatherMoonRegular } from '@fluentui/react-icons';
import { useAppContext } from '../../AppContext';
import { darkTheme, lightTheme } from '@microsoft/mgt-chat';

const availableThemes = [
  {
    key: 'light',
    displayName: 'Light',
    icon: <BrightnessHighRegular />
  },
  {
    key: 'dark',
    displayName: 'Dark',
    icon: <WeatherMoonRegular />
  }
];

export const ThemeSwitcher = () => {
  const [selectedTheme, setSelectedTheme] = React.useState<any>(availableThemes[0]);
  const appContext = useAppContext();

  const onThemeChanged = (theme: any) => {
    setSelectedTheme(theme);
    // Applies the theme to the Fluent UI components
    switch (theme.key) {
      case 'dark':
        appContext.setState({
          ...appContext.state,
          theme: { key: theme.key, fluentTheme: webDarkTheme, chatTheme: darkTheme }
        });
        break;
      case 'light':
      default:
        appContext.setState({
          ...appContext.state,
          theme: { key: theme.key, fluentTheme: webLightTheme, chatTheme: lightTheme }
        });
    }
  };

  return (
    <Menu>
      <MenuTrigger>
        <MenuButton icon={selectedTheme.icon}>{selectedTheme.displayName}</MenuButton>
      </MenuTrigger>

      <MenuPopover>
        <MenuList>
          {availableThemes.map(theme => (
            <MenuItem icon={theme.icon} key={theme.key} onClick={() => onThemeChanged(theme)}>
              {theme.displayName}
            </MenuItem>
          ))}
        </MenuList>
      </MenuPopover>
    </Menu>
  );
};
