import { SearchResults } from '@microsoft/mgt-react/dist/es6/generated/react-preview';
import * as React from 'react';
import { IResultsProps } from './IResultsProps';

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
            ></SearchResults>
          )}
          <SearchResults
            entityTypes={['driveItem', 'listItem', 'site']}
            queryString={props.searchTerm}
            scopes={['Files.Read.All', 'Sites.Read.All']}
            fetchThumbnail={true}
          ></SearchResults>
        </>
      )}
    </>
  );
};
