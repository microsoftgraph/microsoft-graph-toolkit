import { SearchResults } from '@microsoft/mgt-react';
import * as React from 'react';
import { IResultsProps } from './IResultsProps';

export const ExternalItemsResults: React.FunctionComponent<IResultsProps> = (props: IResultsProps) => {
  return (
    <>
      {props.searchTerm && (
        <SearchResults
          entityTypes={['externalItem']}
          contentSources={['/external/connections/contosoBlogPosts']}
          queryString={props.searchTerm}
          scopes={['ExternalItem.Read.All']}
          version="beta"
        ></SearchResults>
      )}
    </>
  );
};
