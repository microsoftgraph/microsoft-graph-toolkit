import {
  Button,
  DataGrid,
  DataGridBody,
  DataGridCell,
  DataGridHeader,
  DataGridHeaderCell,
  DataGridRow,
  TableCellLayout,
  TableColumnDefinition,
  createTableColumn
} from '@fluentui/react-components';
import { SettingsRegular, StarRegular } from '@fluentui/react-icons';
import { Get, MgtTemplateProps, Person, PersonCardInteraction, Providers, ViewType } from '@microsoft/mgt-react';
import './Incidents.css';
import React from 'react';

export interface IIndicentsProps {}

export function Incidents(props: IIndicentsProps) {
  return (
    <Get resource={`sites/root/lists/${process.env.REACT_APP_INCIDENTS_LIST_ID!}/items?$expand=fields`}>
      <DataGridTemplate template="default"></DataGridTemplate>
    </Get>
  );
}

const DataGridTemplate = (props: MgtTemplateProps) => {
  const [incidents] = React.useState<any[]>(props.dataContext.value);

  const columns: TableColumnDefinition<any>[] = [
    createTableColumn<any>({
      columnId: 'title',
      renderHeaderCell: () => {
        return 'Title';
      },
      renderCell: item => {
        return <TableCellLayout>{item.fields.Title}</TableCellLayout>;
      }
    }),
    createTableColumn<any>({
      columnId: 'status',
      renderHeaderCell: () => {
        return 'Status';
      },
      renderCell: item => {
        return <TableCellLayout>{item.fields.Status}</TableCellLayout>;
      }
    }),
    createTableColumn<any>({
      columnId: 'priority',
      renderHeaderCell: () => {
        return 'Priority';
      },
      renderCell: item => {
        return <TableCellLayout>{item.fields.Priority}</TableCellLayout>;
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
            <Get
              resource={`sites/root/lists/${process.env.REACT_APP_USER_INFORMATION_LIST_ID!}/items/${
                item.fields.IssueloggedbyLookupId
              }`}
            >
              <PersonFromUserFormationListTemplate template="default"></PersonFromUserFormationListTemplate>
            </Get>
          </TableCellLayout>
        );
      }
    }),
    createTableColumn<any>({
      columnId: 'actions',
      renderHeaderCell: () => {
        return 'Actions';
      },
      renderCell: item => {
        return (
          <TableCellLayout>
            <Button as="a" icon={<SettingsRegular />} size="small" href={`#/incident/${item.id}`} />
          </TableCellLayout>
        );
      }
    })
  ];

  return (
    <DataGrid columns={columns} items={incidents}>
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
