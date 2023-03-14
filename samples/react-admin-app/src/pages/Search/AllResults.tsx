import { SearchBox, SearchResults } from '@microsoft/mgt-react';
import * as React from 'react';

export const AllResults: React.FunctionComponent = () => {
  const [searchTerm, setSearchTerm] = React.useState<string>('');

  return (
    <>
      <SearchBox searchTermChanged={e => setSearchTerm(e.detail)}></SearchBox>
      <SearchResults entityTypes={['driveItem', 'listItem', 'site']} queryString={searchTerm}></SearchResults>
    </>
  );
};
