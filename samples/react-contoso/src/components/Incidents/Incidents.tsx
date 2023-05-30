import {
  DataGrid,
  DataGridBody,
  DataGridCell,
  DataGridHeader,
  DataGridHeaderCell,
  DataGridRow,
  FluentProvider,
  Menu,
  MenuButton,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  TableCellLayout,
  TableColumnDefinition,
  Toolbar,
  ToolbarButton,
  ToolbarGroup,
  createTableColumn,
  makeStyles,
  webLightTheme
} from '@fluentui/react-components';

import { Skeleton, SkeletonItem } from '@fluentui/react-components/unstable';
import { AddRegular, ContentViewRegular, CheckmarkRegular, ListRegular } from '@fluentui/react-icons';
import { Get, MgtTemplateProps, Person, PersonCardInteraction, Providers, ViewType } from '@microsoft/mgt-react';
import './Incidents.css';
import React, { useRef } from 'react';

export interface IIndicentsProps {}
const useStyles = makeStyles({
  toolbar: {
    justifyContent: 'space-between'
  }
});

const getColumns = (shimmered: boolean): TableColumnDefinition<any>[] => {
  const columns: TableColumnDefinition<any>[] = [
    createTableColumn<any>({
      columnId: 'title',
      renderHeaderCell: () => {
        return 'Title';
      },
      renderCell: item => {
        return (
          <TableCellLayout>
            {shimmered ? <SkeletonItem shape="rectangle" style={{ width: '120px' }} /> : item.fields.Title}
          </TableCellLayout>
        );
      }
    }),
    createTableColumn<any>({
      columnId: 'status',
      renderHeaderCell: () => {
        return 'Status';
      },
      renderCell: item => {
        return (
          <TableCellLayout>
            {shimmered ? <SkeletonItem shape="rectangle" style={{ width: '120px' }} /> : item.fields.Status}
          </TableCellLayout>
        );
      }
    }),
    createTableColumn<any>({
      columnId: 'priority',
      renderHeaderCell: () => {
        return 'Priority';
      },
      renderCell: item => {
        return (
          <TableCellLayout>
            {shimmered ? <SkeletonItem shape="rectangle" style={{ width: '120px' }} /> : item.fields.Priority}
          </TableCellLayout>
        );
      }
    }),
    createTableColumn<any>({
      columnId: 'requestedBy',
      renderHeaderCell: () => {
        return 'Requested By';
      },
      renderCell: item => {
        return (
          <TableCellLayout>
            {shimmered ? (
              <div
                style={{
                  display: 'grid',
                  alignItems: 'center',
                  position: 'relative',
                  gridTemplateColumns: 'min-content 80%',
                  gap: '10px'
                }}
              >
                <SkeletonItem shape="circle" size={32} />
                <SkeletonItem style={{ width: '120px' }} />
              </div>
            ) : (
              <Get
                resource={`sites/root/lists/${process.env.REACT_APP_USER_INFORMATION_LIST_ID!}/items/${
                  item.fields.IssueloggedbyLookupId
                }`}
              >
                <PersonFromUserFormationListTemplate template="default"></PersonFromUserFormationListTemplate>
              </Get>
            )}
          </TableCellLayout>
        );
      }
    })
  ];

  return columns;
};
export function Incidents(props: IIndicentsProps) {
  return (
    <Get resource={`sites/root/lists/${process.env.REACT_APP_INCIDENTS_LIST_ID!}/items?$expand=fields`}>
      <DataGridTemplate template="default"></DataGridTemplate>
      <LoadingTemplate template="loading"></LoadingTemplate>
    </Get>
  );
}

const DataGridTemplate = (props: MgtTemplateProps) => {
  const styles = useStyles();
  //const appContext = useAppContext();
  const [incidents, setIncidents] = React.useState<any[]>(props.dataContext.value);
  const [viewIncidents, setViewIncidents] = React.useState<any[]>(incidents);

  const [selectedIncident, setSelectedIncident] = React.useState<any>(null);
  const [selectedView, setSelectedView] = React.useState<string>('All Incidents');
  const toolbarReference = useRef(null);

  const onSelectionChange = (e: any, data: any) => {
    const [selectedItem] = data.selectedItems;
    const incident = incidents.find(i => i.id === selectedItem);
    setSelectedIncident(incident);
  };

  const onViewMenuClick = async (view: string) => {
    switch (view) {
      case 'My Incidents':
        setSelectedView(view);
        const me = await Providers.me();
        const currentUserInformation = await Providers.globalProvider.graph.client
          .api(
            `/sites/root/lists/${process.env
              .REACT_APP_USER_INFORMATION_LIST_ID!}/items?$expand=fields&$filter=fields/EMail eq '${
              me.userPrincipalName
            }'`
          )
          .headers({ Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' })
          .get();
        setViewIncidents(incidents.filter(i => i.fields.Assignedto0LookupId === currentUserInformation.value[0].id));
        break;
      case 'All Incidents':
        setSelectedView(view);
        setViewIncidents(incidents);
        break;
    }
  };

  const onComplete = async () => {
    let updatedIncidents = [...incidents];
    await Providers.globalProvider.graph.client
      .api(`/sites/root/lists/${process.env.REACT_APP_INCIDENTS_LIST_ID!}/items/${selectedIncident.id}`)
      .patch({
        fields: {
          Status: 'Completed'
        }
      });

    const updatedIncidentIndex = updatedIncidents.findIndex(i => i.id === selectedIncident.id);
    const updatedIncident = updatedIncidents[updatedIncidentIndex];
    updatedIncidents[updatedIncidentIndex] = {
      ...updatedIncident,
      fields: {
        ...updatedIncident.fields,
        Status: 'Completed'
      }
    };

    setIncidents(updatedIncidents);
  };

  return (
    <div>
      <Toolbar className={styles.toolbar} ref={toolbarReference}>
        <ToolbarGroup role="presentation">
          <ToolbarButton icon={<AddRegular />} appearance="primary">
            New Incident
          </ToolbarButton>
          <ToolbarButton
            icon={<ContentViewRegular />}
            disabled={!selectedIncident}
            as="a"
            href={`/incident/${selectedIncident?.id}`}
          >
            View
          </ToolbarButton>
          <ToolbarButton
            icon={<CheckmarkRegular />}
            disabled={!selectedIncident || selectedIncident?.fields.Status === 'Completed'}
            onClick={onComplete}
          >
            Complete
          </ToolbarButton>
        </ToolbarGroup>
        <ToolbarGroup role="presentation">
          <FluentProvider theme={webLightTheme}>
            <Menu>
              <MenuTrigger>
                <MenuButton icon={<ListRegular />}>{selectedView}</MenuButton>
              </MenuTrigger>

              <MenuPopover>
                <MenuList>
                  <MenuItem onClick={() => onViewMenuClick('My Incidents')}>My Incidents</MenuItem>
                  <MenuItem onClick={() => onViewMenuClick('All Incidents')}>All Incidents</MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
          </FluentProvider>
        </ToolbarGroup>
      </Toolbar>
      <DataGrid
        columns={getColumns(false)}
        items={viewIncidents}
        selectionMode="single"
        onSelectionChange={onSelectionChange}
        getRowId={item => item.id}
      >
        <DataGridHeader>
          <DataGridRow>
            {({ renderHeaderCell }) => <DataGridHeaderCell>{renderHeaderCell()}</DataGridHeaderCell>}
          </DataGridRow>
        </DataGridHeader>
        <DataGridBody<any>>
          {({ item, rowId }) => (
            <DataGridRow<any> key={rowId}>
              {({ renderCell }) => <DataGridCell>{renderCell(item)}</DataGridCell>}
            </DataGridRow>
          )}
        </DataGridBody>
      </DataGrid>
    </div>
  );
};

const PersonFromUserFormationListTemplate = (props: MgtTemplateProps) => {
  return (
    <Person
      personQuery={props.dataContext.fields.EMail}
      view={ViewType.oneline}
      personCardInteraction={PersonCardInteraction.hover}
    ></Person>
  );
};

const LoadingTemplate = (props: MgtTemplateProps) => {
  return (
    <DataGrid columns={getColumns(true)} items={[...Array<number>(10)]}>
      <DataGridHeader>
        <DataGridRow>
          {({ renderHeaderCell }) => <DataGridHeaderCell>{renderHeaderCell()}</DataGridHeaderCell>}
        </DataGridRow>
      </DataGridHeader>
      <DataGridBody<any>>
        {({ item, rowId }) => (
          <Skeleton>
            <DataGridRow<any> key={rowId}>
              {({ renderCell }) => <DataGridCell>{renderCell(item)}</DataGridCell>}
            </DataGridRow>
          </Skeleton>
        )}
      </DataGridBody>
    </DataGrid>
  );
};
