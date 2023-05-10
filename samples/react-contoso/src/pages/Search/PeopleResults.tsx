import { SearchBox, SearchResults } from '@microsoft/mgt-react';
import * as React from 'react';

export const PeopleResults: React.FunctionComponent = () => {
  const [searchTerm, setSearchTerm] = React.useState<string>('');

  return (
    <>
      <SearchBox searchTermChanged={e => setSearchTerm(e.detail)}></SearchBox>
      <SearchResults entityTypes={['people']} size={20} queryString={searchTerm} version={'beta'}></SearchResults>
    </>
  );
};
