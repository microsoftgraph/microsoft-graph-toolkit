import { SearchResults } from '@microsoft/mgt-react';
import * as React from 'react';
import { IResultsProps } from './IResultsProps';

export const PeopleResults: React.FunctionComponent<IResultsProps> = (props: IResultsProps) => {
  return (
    <>
      <SearchResults entityTypes={['people']} size={20} queryString={props.searchTerm} version={'beta'}></SearchResults>
    </>
  );
};
