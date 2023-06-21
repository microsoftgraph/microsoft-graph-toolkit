import * as React from 'react';
import { useParams } from 'react-router-dom';
import { PageHeader } from '../../components/PageHeader';
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
import { FileList, Providers } from '@microsoft/mgt-react';
import { Tasks } from '@microsoft/mgt-react';

const useStyles = makeStyles({
  panels: {
    ...shorthands.padding('10px'),
    maxHeight: '100vh',
    overflowY: 'auto'
  },

  dialog: {
    display: 'block'
  }
});

export const Incident: React.FunctionComponent = () => {
  const styles = useStyles();
  const { id } = useParams<{ id: string }>();

  const [selectedTab, setSelectedTab] = React.useState<TabValue>('files');
  const [incident, setIncident] = React.useState<any>();
  const [driveId, setDriveId] = React.useState<string>();
  const [planId, setPlanId] = React.useState<string>();
  const [isLoading, setIsLoading] = React.useState<boolean>(true);

  const onTabSelect = (event: SelectTabEvent, data: SelectTabData) => {
    setSelectedTab(data.value);
  };

  React.useEffect(() => {
    const fetchIncident = async () => {
      const incident = await Providers.globalProvider.graph.client
        .api(`sites/root/lists/${process.env.REACT_APP_INCIDENTS_LIST_ID!}/items/${id}`)
        .get();
      setIncident(incident);
      return incident;
    };

    const initializeDrive = async () => {
      const drive = await Providers.globalProvider.graph.client
        .api(`groups/${process.env.REACT_APP_SUPPORT_GROUP_ID!}/drive`)
        .get();

      setDriveId(drive.id);
    };

    const initializeChannel = async (incident: any) => {
      if (!incident.fields.ChannelId) {
        let channel: any = null;
        const channels = await Providers.globalProvider.graph.client
          .api(
            `teams/${process.env.REACT_APP_SUPPORT_GROUP_ID!}/channels?$filter=displayName eq '${
              incident.fields.Title
            }'`
          )
          .get();

        if (!channels.value.length) {
          channel = await Providers.globalProvider.graph.client
            .api(`teams/${process.env.REACT_APP_SUPPORT_GROUP_ID!}/channels`)
            .post({
              displayName: incident.fields.Title,
              description: incident.fields.Description,
              membershipType: 'standard'
            });
        } else {
          channel = channels.value[0];
        }

        if (channel) {
          await Providers.globalProvider.graph.client
            .api(`sites/root/lists//${process.env.REACT_APP_INCIDENTS_LIST_ID!}/items/${id}`)
            .patch({
              fields: {
                ChannelId: channel.id
              }
            });
        }
      }
    };

    const initializeTasks = async (incident: any) => {
      if (!incident.fields.PlanId) {
        let plan: any = null;

        const plans = await Providers.globalProvider.graph.client
          .api(`groups/${process.env.REACT_APP_SUPPORT_GROUP_ID!}/planner/plans`)
          .get();

        const foundPlan = plans.value.find(p => p.title === incident.fields.Title);

        if (!foundPlan) {
          plan = await Providers.globalProvider.graph.client.api(`planner/plans`).post({
            container: {
              url: `https://graph.microsoft.com/v1.0/groups/${process.env.REACT_APP_SUPPORT_GROUP_ID!}`
            },
            title: incident.fields.Title
          });
        } else {
          plan = foundPlan;
        }

        await Providers.globalProvider.graph.client
          .api(`sites/root/lists//${process.env.REACT_APP_INCIDENTS_LIST_ID!}/items/${id}`)
          .patch({
            fields: {
              PlanId: plan.id
            }
          });

        setPlanId(plan.id);
      } else {
        setPlanId(incident.fields.PlanId);
      }
    };

    const initializeIncident = async () => {
      const currentIncident = await fetchIncident();
      await initializeDrive();
      await initializeChannel(currentIncident);
      await initializeTasks(currentIncident);

      setIsLoading(false);
    };

    initializeIncident();
  }, [id]);

  return (
    <>
      {!isLoading && incident && (
        <>
          <PageHeader title={incident.fields.Title} description={incident.fields.Description}></PageHeader>

          <TabList selectedValue={selectedTab} onTabSelect={onTabSelect}>
            <Tab value="files">Files</Tab>
            <Tab value="tasks">Tasks</Tab>
          </TabList>

          <div className={styles.panels}>
            {selectedTab === 'files' && driveId && (
              <FileList driveId={driveId} itemPath={incident.fields.Title} pageSize={100} />
            )}
            {selectedTab === 'tasks' && planId && <Tasks targetId={planId} />}
          </div>
        </>
      )}
      {isLoading && <Loading message="Loading incident..."></Loading>}
    </>
  );
};
