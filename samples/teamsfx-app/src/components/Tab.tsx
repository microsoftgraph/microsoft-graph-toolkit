import { useContext } from 'react';
import { TeamsFxContext } from './Context';
import React from 'react';
import { Agenda, Person, applyTheme } from '@microsoft/mgt-react';
import { Button } from '@fluentui/react-components';

import { Providers, ProviderState } from '@microsoft/mgt-react';
import { TeamsFxProvider } from '@microsoft/mgt-teamsfx-provider';
import { TeamsUserCredential, TeamsUserCredentialAuthConfig } from '@microsoft/teamsfx';

const authConfig: TeamsUserCredentialAuthConfig = {
  clientId: process.env.REACT_APP_CLIENT_ID!,
  initiateLoginEndpoint: process.env.REACT_APP_START_LOGIN_PAGE_URL!
};

const scopes = ['User.Read', 'Calendars.Read'];
const credential = new TeamsUserCredential(authConfig);
const provider = new TeamsFxProvider(credential, scopes);
Providers.globalProvider = provider;

export default function Tab() {
  const { themeString } = useContext(TeamsFxContext);
  const [loading, setLoading] = React.useState<boolean>(false);
  const [consentNeeded, setConsentNeeded] = React.useState<boolean>(false);

  React.useEffect(() => {
    const init = async () => {
      try {
        await credential.getToken(scopes);
        Providers.globalProvider.setState(ProviderState.SignedIn);
      } catch (error) {
        setConsentNeeded(true);
      }
    };

    init();
  }, []);

  const consent = async () => {
    setLoading(true);
    await credential.login(scopes);
    Providers.globalProvider.setState(ProviderState.SignedIn);
    setLoading(false);
    setConsentNeeded(false);
  };

  React.useEffect(() => {
    applyTheme(themeString === 'default' ? 'light' : 'dark');
  }, [themeString]);

  return (
    <div>
      {consentNeeded && (
        <>
          <p>Click below to authorize button to grant permission to using Microsoft Graph.</p>
          <Button appearance="primary" disabled={loading} onClick={consent}>
            Authorize
          </Button>
        </>
      )}

      {!consentNeeded && (
        <>
          <Person personQuery="me" />
          <Agenda></Agenda>
        </>
      )}
    </div>
  );
}
