import { SearchBox, SearchResults } from '@microsoft/mgt-react';
import * as React from 'react';
import { PageHeader } from '../components/PageHeader/PageHeader';

export const SearchCenter: React.FunctionComponent = () => {
  const [searchTerm, setSearchTerm] = React.useState<string>('');

  return (
    <>
      <PageHeader
        title={'Search'}
        description={'Use this Search Center to test Microsot Graph Toolkit search components capabilities'}
      ></PageHeader>
      <SearchBox searchTermChanged={e => setSearchTerm(e.detail)}></SearchBox>
      <SearchResults queryString={searchTerm}></SearchResults>
    </>
  );
};
