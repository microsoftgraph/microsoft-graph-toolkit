import * as React from 'react';
import { Login, SearchBox, useIsSignedIn } from '@microsoft/mgt-react';
import { PACKAGE_VERSION } from '@microsoft/mgt-element';
import { SimpleLogin } from './SimpleLogin';
import { useNavigate } from 'react-router-dom';
import { ThemeSwitcher } from './ThemeSwitcher';
import { useAppContext } from '../AppContext';
import { Label, makeStyles, mergeClasses, shorthands, tokens, InfoLabel } from '@fluentui/react-components';
import { GridDotsRegular } from '@fluentui/react-icons';
import { useLocation } from 'react-router-dom';

const useStyles = makeStyles({
  root: {
    '--login-signed-in-hover-background': 'transparent'
  },

  header: {
    display: 'flex',
    flexDirection: 'row',
    flexWrap: 'nowrap',
    width: '100%',
    height: 'auto',
    boxSizing: 'border-box',
    alignItems: 'center',
    backgroundColor: tokens.colorNeutralForeground1Static,
    zIndex: '1',
    minHeight: '50px'
  },

  name: {
    color: tokens.colorNeutralForegroundOnBrand,
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase400,
    paddingLeft: '12px'
  },

  wafffleIcon: {
    fontSize: tokens.fontSizeBase600,
    paddingTop: '5px',
    color: tokens.colorNeutralForegroundOnBrand
  },

  waffle: {
    display: 'flex',
    width: 'auto',
    height: 'auto',
    boxSizing: 'border-box',
    flexGrow: '1',
    alignItems: 'center',
    minWidth: 'max-content',
    paddingLeft: '8px'
  },

  waffleLogo: {
    ...shorthands.flex('none')
  },

  waffleTitle: {
    ...shorthands.flex('auto'),
    display: 'flex',
    alignItems: 'center'
  },

  search: {
    display: 'flex',
    width: '100%',
    height: 'auto',
    boxSizing: 'border-box',
    flexGrow: '1',
    justifyContent: 'center'
  },

  searchBox: {
    minWidth: '320px',
    maxWidth: '468px',
    paddingRight: '1em',
    paddingLeft: '1em',
    width: '100%'
  },

  infoIcon: {
    color: tokens.colorNeutralForegroundOnBrand,
    ':hover': {
      color: tokens.colorNeutralForegroundOnBrand,
      ':active': {
        color: tokens.colorNeutralForegroundOnBrand
      }
    }
  },

  login: {
    display: 'flex',
    width: 'auto',
    height: 'auto',
    boxSizing: 'border-box',
    flexGrow: '1',
    alignItems: 'center',
    minWidth: 'max-content'
  },

  signedOut: {
    ...shorthands.padding('4px', '8px')
  },

  signedIn: {
    ...shorthands.padding('0px', '10px')
  }
});

const HeaderComponent: React.FunctionComponent = () => {
  const styles = useStyles();
  const [isSignedIn] = useIsSignedIn();
  const appContext = useAppContext();
  const setAppContext = appContext.setState;
  const location = useLocation();
  const navigate = useNavigate();

  const onSearchTermChanged = (e: CustomEvent) => {
    if (!(e.detail === '' && appContext.state.searchTerm === '*') && e.detail !== appContext.state.searchTerm) {
      appContext.setState({ ...appContext.state, searchTerm: e.detail === '' ? '*' : e.detail });

      if (e.detail === '') {
        navigate('/search');
      } else {
        navigate('/search?q=' + e.detail);
      }
    }
  };

  React.useLayoutEffect(() => {
    if (location.pathname === '/search') {
      const searchTerm = decodeURI(location.search.replace('?q=', ''));
      setAppContext(previous => {
        return { ...previous, searchTerm: searchTerm === '' ? '*' : searchTerm };
      });
    }
  }, [location, setAppContext]);

  return (
    <div className={styles.header}>
      <div className={styles.waffle}>
        <div className={styles.waffleLogo}>
          <a href={'https://www.office.com/apps?auth=2'} target="_blank" rel="noreferrer">
            <GridDotsRegular className={styles.wafffleIcon} />
          </a>
        </div>

        <div className={styles.waffleTitle}>
          <Label className={styles.name}>{import.meta.env.VITE_SITE_NAME} </Label>
          <InfoLabel className={styles.infoIcon} size="medium" info={<>Using the Graph Toolkit v{PACKAGE_VERSION}</>} />
        </div>
      </div>
      <div className={styles.search}>
        <SearchBox
          className={styles.searchBox}
          searchTermChanged={onSearchTermChanged}
          searchTerm={appContext.state.searchTerm !== '*' ? appContext.state.searchTerm : ''}
        ></SearchBox>
      </div>

      <div className={styles.login}>
        <ThemeSwitcher />
        <div className={mergeClasses(!isSignedIn ? styles.signedOut : styles.signedIn, styles.root)}>
          <Login>
            <SimpleLogin template="signed-in-button-content" />
          </Login>
        </div>
      </div>
    </div>
  );
};

export const Header = React.memo(HeaderComponent);
