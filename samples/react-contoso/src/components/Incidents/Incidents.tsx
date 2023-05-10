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
import { SettingsRegular, StarRegular, StarFilled } from '@fluentui/react-icons';
import { Person, PersonCardInteraction, ViewType } from '@microsoft/mgt-react';
import './Incidents.css';
import React from 'react';

export interface IIndicentsProps {}

export function Incidents(props: IIndicentsProps) {
  const [incidents, setIncidents] = React.useState<any[]>([]);

  React.useEffect(() => {
    const fetchData = async () => {
      const response = await fetch('incidents.json');
      const data = await response.json();
      setIncidents(data);
    };

    fetchData();
  }, []);

  const columns: TableColumnDefinition<any>[] = [
    createTableColumn<any>({
      columnId: 'title',
      renderHeaderCell: () => {
        return 'Title';
      },
      renderCell: item => {
        return <TableCellLayout>{item.title}</TableCellLayout>;
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
            <Person
              userId={item.requestedBy}
              view={ViewType.oneline}
              personCardInteraction={PersonCardInteraction.hover}
            ></Person>
          </TableCellLayout>
        );
      }
    }),
    createTableColumn<any>({
      columnId: 'description',
      renderHeaderCell: () => {
        return 'Title';
      },
      renderCell: item => {
        return <TableCellLayout>{item.description}</TableCellLayout>;
      }
    }),
    createTableColumn<any>({
      columnId: 'status',
      renderHeaderCell: () => {
        return 'Status';
      },
      renderCell: item => {
        return <TableCellLayout>{item.status}</TableCellLayout>;
      }
    }),
    createTableColumn<any>({
      columnId: 'priority',
      renderHeaderCell: () => {
        return 'Priority';
      },
      renderCell: item => {
        return <TableCellLayout>{item.priority}</TableCellLayout>;
      }
    }),
    createTableColumn<any>({
      columnId: 'assignedTo',
      renderHeaderCell: () => {
        return 'Assigned To';
      },
      renderCell: item => {
        return (
          <TableCellLayout>
            <Person
              userId={item.assignedTo}
              view={ViewType.oneline}
              personCardInteraction={PersonCardInteraction.hover}
            ></Person>
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
            <Button icon={<StarRegular />} size="small" />
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
}
