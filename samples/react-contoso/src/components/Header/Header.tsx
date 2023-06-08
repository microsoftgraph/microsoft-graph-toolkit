import * as React from 'react';
import { Login } from '@microsoft/mgt-react';
import { SearchBox } from '@microsoft/mgt-react/dist/es6/generated/react-preview';
import { PACKAGE_VERSION } from '@microsoft/mgt-element';
import { InfoButton } from '@fluentui/react-components/unstable';
import { SimpleLogin } from '../SimpleLogin/SimpleLogin';
import { useIsSignedIn } from '../../hooks/useIsSignedIn';
import './Header.css';
import { useHistory } from 'react-router-dom';
import { ThemeSwitcher } from '../ThemeSwitcher/ThemeSwitcher';
import { useAppContext } from '../../AppContext';
import { Label, makeStyles, shorthands, tokens } from '@fluentui/react-components';
import { GridDotsRegular } from '@fluentui/react-icons';

const useStyles = makeStyles({
  name: {
    color: tokens.colorNeutralForegroundOnBrand,
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase400,
    paddingLeft: '8px'
  },
  wafffleIcon: {
    fontSize: tokens.fontSizeBase600,
    ...shorthands.margin('0 10px'),
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
  }
});

const HeaderComponent: React.FunctionComponent = () => {
  const styles = useStyles();
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
      <div className={styles.waffle}>
        <div className={styles.waffleLogo}>
          <a href={'https://myapps.microsoft.com'} target="_blank" rel="noreferrer">
            <GridDotsRegular className={styles.wafffleIcon} />
          </a>
        </div>

        <div className={styles.waffleTitle}>
          <Label className={styles.name}>{process.env.REACT_APP_SITE_NAME} </Label>
          <InfoButton className={styles.infoIcon} size="medium" info={<>{PACKAGE_VERSION}</>} />
        </div>
      </div>
      <div className={styles.search}>
        <SearchBox className={styles.searchBox} searchTermChanged={onSearchTermChanged}></SearchBox>
      </div>

      <div className="login">
        <ThemeSwitcher />
        <Login className={!isSignedIn ? 'signed-out' : 'signed-in'}>
          <SimpleLogin template="signed-in-button-content" />
        </Login>
      </div>
    </div>
  );
};

export const Header = React.memo(HeaderComponent);
