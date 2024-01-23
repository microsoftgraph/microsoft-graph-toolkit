import { SearchResults } from '@microsoft/mgt-react';
import * as React from 'react';
import { IResultsProps } from './IResultsProps';
import { MgtTemplateProps, Person, File } from '@microsoft/mgt-react';
import {
  DataGrid,
  DataGridBody,
  DataGridCell,
  DataGridHeader,
  DataGridHeaderCell,
  DataGridRow,
  SkeletonItem,
  TableCellLayout,
  TableColumnDefinition,
  createTableColumn,
  makeStyles,
  shorthands,
  tokens
} from '@fluentui/react-components';
import { SlideSearchRegular } from '@fluentui/react-icons';

const useStyles = makeStyles({
  container: {
    ...shorthands.gap('16px'),
    display: 'flex',
    flexDirection: 'row',
    flexWrap: 'wrap'
  },
  card: {
    width: '300px',
    height: 'fit-content',
    maxWidth: '100%'
  },
  caption: {
    color: tokens.colorNeutralForeground3
  },
  noDataSearchTerm: {
    fontWeight: tokens.fontWeightSemibold
  },
  emptyContainer: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    flexDirection: 'column',
    height: 'calc(100vh - 300px)'
  },
  fileContainer: {
    display: 'flex'
  },
  fileTitle: {
    paddingLeft: '10px',
    alignSelf: 'center'
  },
  noDataMessage: {
    paddingLeft: '10px'
  },
  noDataIcon: {
    fontSize: '128px'
  },
  row: {
    cursor: 'pointer'
  }
});

const pageSize = 30;

export const FilesResults: React.FunctionComponent<IResultsProps> = (props: IResultsProps) => {
  return (
    <>
      {props.searchTerm && (
        <SearchResults
          entityTypes={['driveItem']}
          queryString={props.searchTerm}
          fetchThumbnail={true}
          queryTemplate="({searchTerms}) ContentTypeId:0x0101*"
          version="beta"
          fields={['createdBy', 'lastModifiedDateTime', 'Title', 'DefaultEncodingURL']}
          size={pageSize}
          cacheEnabled={true}
        >
          <FileTemplate template="default"></FileTemplate>
          <FileTemplate template="loading"></FileTemplate>
          <FileNoDataTemplate template="no-data"></FileNoDataTemplate>
        </SearchResults>
      )}
    </>
  );
};

const getColumns = (shimmered: boolean, styles): TableColumnDefinition<any>[] => {
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
              <div className={styles.fileContainer}>
                <File fileDetails={item.resource} view="image" />
                <span className={styles.fileTitle}>{item.resource.listItem.fields.title}</span>
              </div>
            )}
          </TableCellLayout>
        );
      }
    }),
    createTableColumn<any>({
      columnId: 'modified',
      renderHeaderCell: () => {
        return 'Modified';
      },
      renderCell: item => {
        return (
          <TableCellLayout>
            {shimmered ? (
              <SkeletonItem shape="rectangle" style={{ width: '120px' }} />
            ) : (
              getRelativeDisplayDate(new Date(item.resource.lastModifiedDateTime))
            )}
          </TableCellLayout>
        );
      }
    }),
    createTableColumn<any>({
      columnId: 'owner',
      renderHeaderCell: () => {
        return 'Owner';
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
              <Person personQuery={item.resource.createdBy.user.email} view="oneline" personCardInteraction="hover" />
            )}
          </TableCellLayout>
        );
      }
    })
  ];

  return columns;
};

const FileTemplate = (props: MgtTemplateProps) => {
  const styles = useStyles();
  const [driveItems] = React.useState<any>(props.dataContext.value?.[0]?.hitsContainers[0]?.hits);

  const onRowClick = (item: any) => {
    const url = new URL(item.resource.listItem.fields.defaultEncodingURL);
    url.searchParams.append('Web', '1');
    window.open(url.toString(), '_blank');
  };

  return (
    <div>
      <DataGrid
        columns={getColumns(props.template === 'loading', styles)}
        items={props.template === 'loading' ? [...Array<number>(pageSize)] : driveItems}
      >
        <DataGridHeader>
          <DataGridRow>
            {({ renderHeaderCell }) => <DataGridHeaderCell>{renderHeaderCell()}</DataGridHeaderCell>}
          </DataGridRow>
        </DataGridHeader>
        <DataGridBody<any>>
          {({ item, rowId }) => (
            <DataGridRow<any> key={rowId} className={styles.row} onClick={() => onRowClick(item)}>
              {({ renderCell }) => <DataGridCell>{renderCell(item)}</DataGridCell>}
            </DataGridRow>
          )}
        </DataGridBody>
      </DataGrid>
    </div>
  );
};

const FileNoDataTemplate = (props: MgtTemplateProps) => {
  const styles = useStyles();
  const [searchTerms] = React.useState<string[]>(props.dataContext.value[0]?.searchTerms);

  return (
    <div className={styles.emptyContainer}>
      <div>
        <SlideSearchRegular className={styles.noDataIcon} />
      </div>
      <div className={styles.noDataMessage}>
        We couldn't find any results for <span className={styles.noDataSearchTerm}>{searchTerms.join(' ')}</span>
      </div>
    </div>
  );
};

const getRelativeDisplayDate = (date: Date): string => {
  const now = new Date();

  // Today -> 5:23 PM
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  if (date >= today) {
    return date.toLocaleString('default', {
      hour: 'numeric',
      minute: 'numeric'
    });
  }

  // This week -> Sun 3:04 PM
  const sunday = new Date(today);
  sunday.setDate(now.getDate() - now.getDay());
  if (date >= sunday) {
    return date.toLocaleString('default', {
      hour: 'numeric',
      minute: 'numeric',
      weekday: 'short'
    });
  }

  // Last two week -> Sun 8/2
  const lastTwoWeeks = new Date(sunday);
  lastTwoWeeks.setDate(sunday.getDate() - 7);
  if (date >= lastTwoWeeks) {
    return date.toLocaleString('default', {
      day: 'numeric',
      month: 'numeric',
      weekday: 'short'
    });
  }

  // More than two weeks ago -> 8/1/2020
  return date.toLocaleString('default', {
    day: 'numeric',
    month: 'numeric',
    year: 'numeric'
  });
};
