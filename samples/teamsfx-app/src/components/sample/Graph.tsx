import { Providers, ProviderState } from '@microsoft/mgt-element';
import { TeamsFxProvider } from '@microsoft/mgt-teamsfx-provider';
import { Button } from '@fluentui/react-components';
import React from 'react';
import { Agenda } from '@microsoft/mgt-react';
import { TeamsUserCredential, TeamsUserCredentialAuthConfig } from '@microsoft/teamsfx';

const scopes = ['User.Read', 'Calendars.Read'];

export function Graph() {
  const authConfig: TeamsUserCredentialAuthConfig = {
    clientId: process.env.REACT_APP_CLIENT_ID!,
    initiateLoginEndpoint: process.env.REACT_APP_START_LOGIN_PAGE_URL!
  };

  const [loading, setLoading] = React.useState<boolean>(false);
  const [consentNeeded, setConsentNeeded] = React.useState<boolean>(false);
  const [credential] = React.useState<TeamsUserCredential>(new TeamsUserCredential(authConfig));

  React.useEffect(() => {
    const initAuth = async (scopes: string[]) => {
      const provider = new TeamsFxProvider(credential, scopes);
      Providers.globalProvider = provider;

      try {
        await credential.getToken(scopes);
        Providers.globalProvider.setState(ProviderState.SignedIn);
      } catch (error) {
        setConsentNeeded(true);
      }
    };

    initAuth(scopes);
  }, [credential]);

  const consent = async () => {
    setLoading(true);
    await credential.login(scopes);
    Providers.globalProvider.setState(ProviderState.SignedIn);
    setLoading(false);
    setConsentNeeded(false);
  };

  return (
    <div>
      <div className="section-margin">
        {consentNeeded && (
          <>
            <p>Click below to authorize button to grant permission to using Microsoft Graph.</p>
            <Button appearance="primary" disabled={loading} onClick={consent}>
              Authorize
            </Button>
          </>
        )}

        {!consentNeeded && <Agenda></Agenda>}
      </div>
    </div>
  );
}
