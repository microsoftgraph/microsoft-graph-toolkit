import { SearchBox, SearchResults } from '@microsoft/mgt-react';
import * as React from 'react';

export const Interleaving: React.FunctionComponent = () => {
  const [searchTerm, setSearchTerm] = React.useState<string>('');

  return (
    <>
      <SearchBox searchTermChanged={e => setSearchTerm(e.detail)}></SearchBox>
      <SearchResults
        entityTypes={['externalItem', 'driveItem', 'listItem', 'site']}
        contentSources={['/external/connections/contosoProducts']}
        queryString={searchTerm}
      ></SearchResults>
    </>
  );
};
