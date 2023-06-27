import * as React from 'react';
import {
  Menu,
  MenuButton,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  webDarkTheme,
  webLightTheme,
  teamsLightTheme,
  teamsDarkTheme
} from '@fluentui/react-components';
import { BrightnessHighRegular, WeatherMoonFilled, PeopleTeamRegular, PeopleTeamFilled } from '@fluentui/react-icons';
import { useAppContext } from '../AppContext';

const availableThemes = [
  {
    key: 'light',
    displayName: 'Web Light',
    icon: <BrightnessHighRegular />
  },
  {
    key: 'dark',
    displayName: 'Web Dark',
    icon: <WeatherMoonFilled />
  },
  {
    key: 'teamsLight',
    displayName: 'Teams Light',
    icon: <PeopleTeamRegular />
  },
  {
    key: 'teamsDark',
    displayName: 'Teams Dark',
    icon: <PeopleTeamFilled />
  }
];

export const ThemeSwitcher = () => {
  const [selectedTheme, setSelectedTheme] = React.useState<any>(availableThemes[0]);
  const appContext = useAppContext();

  const onThemeChanged = (theme: any) => {
    setSelectedTheme(theme);
    // Applies the theme to the Fluent UI components
    switch (theme.key) {
      case 'teamsLight':
        appContext.setState({
          ...appContext.state,
          theme: { key: 'light', fluentTheme: teamsLightTheme }
        });
        break;
      case 'teamsDark':
        appContext.setState({
          ...appContext.state,
          theme: { key: 'dark', fluentTheme: teamsDarkTheme }
        });
        break;
      case 'light':
        appContext.setState({
          ...appContext.state,
          theme: { key: theme.key, fluentTheme: webLightTheme }
        });
        break;
      case 'dark':
        appContext.setState({
          ...appContext.state,
          theme: { key: theme.key, fluentTheme: webDarkTheme }
        });
        break;
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
