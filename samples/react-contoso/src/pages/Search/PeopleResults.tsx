import { SearchResults } from '@microsoft/mgt-react';
import * as React from 'react';

export interface IPeopleResultsProps {
  searchTerm: string;
}

export const PeopleResults: React.FunctionComponent<IPeopleResultsProps> = (props: IPeopleResultsProps) => {
  return (
    <>
      <SearchResults entityTypes={['people']} size={20} queryString={props.searchTerm} version={'beta'}></SearchResults>
    </>
  );
};
