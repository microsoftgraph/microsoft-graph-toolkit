import { SearchResults } from '@microsoft/mgt-react/dist/es6/generated/react-preview';
import * as React from 'react';
import { IResultsProps } from './IResultsProps';


export const ExternalItems: React.FunctionComponent<IResultsProps> = (props: IResultsProps) => {
  return (
    <>
      {props.searchTerm && (

          <SearchResults
            entityTypes={['externalItem']}
            contentSources={["/external/connections/contosoBlogPosts"]}
            queryString={props.searchTerm}
            scopes={['ExternalItem.Read.All']}
            version='beta'
          ></SearchResults>

      )}
    </>
  );
};
