import * as React from 'react';
import { useParams } from 'react-router-dom';
import { PageHeader } from '../../components/PageHeader/PageHeader';
import { Loading } from '../../components/Loading';
import {
  SelectTabData,
  SelectTabEvent,
  Tab,
  TabList,
  TabValue,
  shorthands,
  makeStyles
} from '@fluentui/react-components';
import { FileListComposite } from '@microsoft/mgt-react';
import { Chat, allChatScopes } from '@microsoft/mgt-chat';
import { Tasks } from '@microsoft/mgt-react';

const useStyles = makeStyles({
  panels: {
    ...shorthands.padding('10px')
  }
});

/**
 * Object mapping chat operations to the scopes required to perform them
 */
const incidentOperationScopes: Record<string, string[]> = {
  conversation: [...allChatScopes],
  tasks: ['group.readwrite.all'],
  files: ['files.readwrite.all']
};

/**
 * Provides an array of the distinct scopes required for all chat operations
 */
export const allIncidentScopes = Array.from(
  Object.values(incidentOperationScopes).reduce((acc, scopes) => {
    scopes.forEach(s => acc.add(s));
    return acc;
  }, new Set<string>())
);

export const Incident: React.FunctionComponent = () => {
  const styles = useStyles();
  const { id } = useParams();

  const [selectedTab, setSelectedTab] = React.useState<TabValue>('files');
  const [incident, setIncident] = React.useState<any>();
  const [isLoading, setIsLoading] = React.useState<boolean>(true);

  const onTabSelect = (event: SelectTabEvent, data: SelectTabData) => {
    setSelectedTab(data.value);
  };

  React.useEffect(() => {
    const fetchData = async () => {
      const response = await fetch('incidents.json');
      const data = await response.json();
      const foundIncident = data.find((i: any) => i.id == id);
      setIncident(foundIncident);
      setIsLoading(false);
    };

    fetchData();
  }, [id]);

  return (
    <>
      {!isLoading && incident && (
        <>
          <PageHeader title={incident.title} description={incident.description}></PageHeader>

          <TabList selectedValue={selectedTab} onTabSelect={onTabSelect}>
            <Tab value="files">Files</Tab>
            <Tab value="tasks">Tasks</Tab>
            <Tab value="conversation">Conversation</Tab>
          </TabList>

          <div className={styles.panels}>
            {selectedTab === 'files' && (
              <FileListComposite
                enableCommandBar={true}
                breadcrumbRootName="Relevant Documents"
                enableFileUpload={true}
                useGridView={true}
                driveId={incident.driveId}
                itemPath={incident.itemPath}
                pageSize={100}
              />
            )}
            {selectedTab === 'tasks' && incident.planId && <Tasks targetId={incident.planId} />}
            {selectedTab === 'conversation' && incident.conversationId && <Chat chatId={incident.conversationId} />}
          </div>
        </>
      )}
      {isLoading && <Loading message="Loading incident..."></Loading>}
    </>
  );
};
