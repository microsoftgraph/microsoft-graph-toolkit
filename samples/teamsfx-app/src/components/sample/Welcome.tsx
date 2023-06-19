import { useContext } from 'react';
import { Image } from '@fluentui/react-components';
import './Welcome.css';
import { app } from '@microsoft/teams-js';
import { Graph } from './Graph';
import { useData } from '@microsoft/teamsfx-react';
import { TeamsFxContext } from '../Context';

export function Welcome(props: { showFunction?: boolean; environment?: string }) {
  const { environment } = {
    environment: window.location.hostname === 'localhost' ? 'local' : 'azure',
    ...props
  };
  const friendlyEnvironmentName =
    {
      local: 'local environment',
      azure: 'Azure environment'
    }[environment] || 'local environment';

  const { teamsUserCredential } = useContext(TeamsFxContext);
  const { loading, data, error } = useData(async () => {
    if (teamsUserCredential) {
      const userInfo = await teamsUserCredential.getUserInfo();
      return userInfo;
    }
  });
  const userName = loading || error ? '' : data!.displayName;
  const hubName = useData(async () => {
    await app.initialize();
    const context = await app.getContext();
    return context.app.host.name;
  })?.data;
  return (
    <div className="welcome page">
      <div className="narrow page-padding">
        <Image src="hello.png" />
        <h1 className="center">Congratulations{userName ? ', ' + userName : ''}!</h1>
        {hubName && <p className="center">Your app is running in {hubName}</p>}
        <p className="center">Your app is running in your {friendlyEnvironmentName}</p>

        <div>
          <Graph />
        </div>
      </div>
    </div>
  );
}
