import * as React from 'react';
import { IResultsProps } from './IResultsProps';
import { MgtTemplateProps, SearchResults } from '@microsoft/mgt-react';

export const AllResults: React.FunctionComponent<IResultsProps> = (props: IResultsProps) => {
  return (
    <>
      {props.searchTerm && (
        <>
          {props.searchTerm !== '*' && (
            <SearchResults
              entityTypes={['bookmark']}
              queryString={props.searchTerm}
              version="beta"
              size={1}
              scopes={['Bookmark.Read.All']}
            >
              <NoDataTemplate template="no-data"></NoDataTemplate>
            </SearchResults>
          )}
          <SearchResults
            entityTypes={['driveItem', 'listItem', 'site']}
            queryString={props.searchTerm}
            scopes={['Files.Read.All', 'Files.ReadWrite.All', 'Sites.Read.All']}
            fetchThumbnail={true}
          ></SearchResults>
        </>
      )}
    </>
  );
};

const NoDataTemplate = (props: MgtTemplateProps) => {
  return <></>;
};
