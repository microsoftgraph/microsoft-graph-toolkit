import * as React from 'react';
import { Stack, ILinkStyleProps, ILinkStyles, ITheme, IStackProps, FontIcon, Label } from '@fluentui/react';
import { mergeStyles, Stylesheet } from '@fluentui/merge-styles';
import { Login } from '@microsoft/mgt-react';
import { SimpleLogin } from '../SimpleLogin/SimpleLogin';
import { useIsSignedIn } from '../../hooks/useIsSignedIn';
import './Header.css';

export interface IHeaderProps {
  //onGenerate: (event: React.MouseEvent<HTMLButtonElement>) => void;
}

const pipeFabricStyles = (p: ILinkStyleProps): ILinkStyles => ({
  root: {
    textDecoration: 'none',
    color: '#ffffff',
    fontWeight: '600',
    fontSize: p.theme.fonts.mediumPlus.fontSize,
    paddingLeft: '8px',
    focus: {
      textDecorationColor: '#ffffff',
      textDecorationStyle: 'underline'
    }
  }
});

const headerStackStyles = (p: IStackProps, theme: ITheme) => ({
  root: {
    backgroundColor: theme.semanticColors.bodyBackground,
    minHeight: 50
  }
});

const headerStyles = mergeStyles({
  backgroundColor: '#334A5F',
  position: 'fixed',
  top: 0,
  left: 0,
  width: '100%',
  zIndex: 1
});

const waffleIconClass = mergeStyles({
  fontSize: 24,
  margin: '0 10px',
  paddingTop: '5px',
  color: 'var(--color-brand-iconColor)'
});

const HeaderComponent: React.FunctionComponent<IHeaderProps> = (props: IHeaderProps) => {
  const [isSignedIn] = useIsSignedIn();

  return (
    <Stack horizontal verticalAlign="center" grow={0} styles={headerStackStyles} className={headerStyles}>
      <Stack horizontal grow={1} verticalAlign="center">
        <a href={'https://myapps.microsoft.com'} target="_blank" rel="noreferrer">
          <FontIcon iconName="Waffle" className={waffleIconClass} />
        </a>

        <Label styles={pipeFabricStyles}>{process.env.REACT_APP_SITE_NAME}</Label>
      </Stack>
      <Login className={!isSignedIn ? 'signed-out' : ''}>
        <SimpleLogin template="signed-in-button-content" />
      </Login>
    </Stack>
  );
};

export const Header = React.memo(HeaderComponent);
