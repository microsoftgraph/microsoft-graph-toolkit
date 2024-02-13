import {
  DataGrid,
  DataGridBody,
  DataGridCell,
  DataGridHeader,
  DataGridHeaderCell,
  DataGridRow,
  TableCellLayout,
  TableColumnDefinition,
  Toolbar,
  ToolbarButton,
  ToolbarGroup,
  createTableColumn,
  makeStyles
} from '@fluentui/react-components';

import { SkeletonItem } from '@fluentui/react-components';
import { FeedRegular } from '@fluentui/react-icons';
import { Get, MgtTemplateProps, Person } from '@microsoft/mgt-react';
import React from 'react';

export interface IIndicentsProps {}
const useStyles = makeStyles({
  toolbar: {
    justifyContent: 'space-between'
  }
});

const getColumns = (shimmered: boolean): TableColumnDefinition<any>[] => {
  const columns: TableColumnDefinition<any>[] = [
    createTableColumn<any>({
      columnId: 'name',
      renderHeaderCell: () => {
        return 'Name';
      },
      renderCell: item => {
        return (
          <TableCellLayout>
            {shimmered ? (
              <SkeletonItem shape="rectangle" style={{ width: '120px' }} />
            ) : (
              <Person userId={item.id} view="oneline" personCardInteraction="hover" />
            )}
          </TableCellLayout>
        );
      }
    }),
    createTableColumn<any>({
      columnId: 'jobTitle',
      renderHeaderCell: () => {
        return 'Job Title';
      },
      renderCell: item => {
        return (
          <TableCellLayout>
            {shimmered ? <SkeletonItem shape="rectangle" style={{ width: '120px' }} /> : item.jobTitle}
          </TableCellLayout>
        );
      }
    }),
    createTableColumn<any>({
      columnId: 'mobilePhone',
      renderHeaderCell: () => {
        return 'Mobile Phone';
      },
      renderCell: item => {
        return (
          <TableCellLayout>
            {shimmered ? <SkeletonItem shape="rectangle" style={{ width: '120px' }} /> : item.mobilePhone}
          </TableCellLayout>
        );
      }
    }),
    createTableColumn<any>({
      columnId: 'officeLocation',
      renderHeaderCell: () => {
        return 'Office Location';
      },
      renderCell: item => {
        return (
          <TableCellLayout>
            {shimmered ? <SkeletonItem shape="rectangle" style={{ width: '120px' }} /> : item.officeLocation}
          </TableCellLayout>
        );
      }
    })
  ];

  return columns;
};

export function DirectReports(props: IIndicentsProps) {
  return (
    <Get resource={`me/directReports`}>
      <DataGridTemplate template="default"></DataGridTemplate>
      <DataGridTemplate template="loading"></DataGridTemplate>
      <NoDataTemplate template="no-data"></NoDataTemplate>
    </Get>
  );
}

const DataGridTemplate = (props: MgtTemplateProps) => {
  const styles = useStyles();
  const [teams] = React.useState<any[]>(props.dataContext.value);
  const [isLoading] = React.useState<boolean>(props.dataContext && !props.dataContext.value);
  const [selectedTeam, setSelectedTeam] = React.useState<any>(null);

  const onSelectionChange = (e: any, data: any) => {
    const [selectedItem] = data.selectedItems;
    const team = teams.find(i => i.id === selectedItem);
    setSelectedTeam(team);
  };

  return (
    <div>
      <Toolbar className={styles.toolbar}>
        <ToolbarGroup role="presentation">
          <ToolbarButton
            icon={<FeedRegular />}
            disabled={!selectedTeam}
            as="a"
            href={`https://www.office.com/feed?auth=2#/user/${selectedTeam?.id}`}
            target="_blank"
          >
            View feed
          </ToolbarButton>
        </ToolbarGroup>
      </Toolbar>
      <DataGrid
        columns={getColumns(isLoading)}
        items={isLoading ? [...Array<number>(5)] : teams}
        selectionMode="single"
        onSelectionChange={onSelectionChange}
        getRowId={item => (isLoading ? Math.random() : item.id)}
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

const NoDataTemplate = (props: MgtTemplateProps) => {
  return <>You don't have direct reports</>;
};
