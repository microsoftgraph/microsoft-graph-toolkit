import { INavLinkGroup, INavStyles, Nav } from '@fluentui/react';
import * as React from 'react';

export interface ISideNavigationProps {
  items: INavLinkGroup[];
}

export const SideNavigation: React.FunctionComponent<ISideNavigationProps> = props => {
  const { items } = props;

  const navStyles: Partial<INavStyles> = {
    groupContent: {
      animationDuration: '0'
    },
    root: {
      overflowY: 'auto',
      //width: '280px',
      backgroundColor: 'rgb(233, 233, 233)'
    },
    link: {},
    compositeLink: {}
  };

  return <Nav ariaLabel="Navigation" styles={navStyles} groups={items} />;
};
