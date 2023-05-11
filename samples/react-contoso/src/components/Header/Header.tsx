import * as React from 'react';
import { ILinkStyleProps, ILinkStyles, FontIcon, Label } from '@fluentui/react';
import { mergeStyles } from '@fluentui/merge-styles';
import { Login, SearchBox } from '@microsoft/mgt-react';
import { SimpleLogin } from '../SimpleLogin/SimpleLogin';
import { useIsSignedIn } from '../../hooks/useIsSignedIn';
import './Header.css';
import { useAppContext } from '../../hooks/useAppContext';
import { useHistory } from 'react-router-dom';

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

const waffleIconClass = mergeStyles({
  fontSize: 24,
  margin: '0 10px',
  paddingTop: '5px',
  color: 'var(--color-brand-iconColor)'
});

const HeaderComponent: React.FunctionComponent<IHeaderProps> = (props: IHeaderProps) => {
  const [isSignedIn] = useIsSignedIn();
  const appContext = useAppContext();
  const history = useHistory();

  const onSearchTermChanged = (e: CustomEvent) => {
    appContext.setState({ ...appContext.state, searchTerm: e.detail });

    if (e.detail === '') {
      history.push('/home');
    } else {
      history.push('/search');
    }
  };

  return (
    <div className="header">
      <div className="waffle">
        <div className="logo">
          <a href={'https://myapps.microsoft.com'} target="_blank" rel="noreferrer">
            <FontIcon iconName="Waffle" className={waffleIconClass} />
          </a>
        </div>

        <div className="title">
          <Label styles={pipeFabricStyles}>{process.env.REACT_APP_SITE_NAME}</Label>
        </div>
      </div>
      <div className="search">
        <SearchBox className="header-search" searchTermChanged={onSearchTermChanged}></SearchBox>
      </div>

      <div className="login">
        <Login className={!isSignedIn ? 'signed-out' : 'signed-in'}>
          <SimpleLogin template="signed-in-button-content" />
        </Login>
      </div>
    </div>
  );
};

export const Header = React.memo(HeaderComponent);
