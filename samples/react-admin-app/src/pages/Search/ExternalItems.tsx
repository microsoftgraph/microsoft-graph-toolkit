import { SearchBox, SearchResults } from '@microsoft/mgt-react';
import * as React from 'react';

export const ExternalItems: React.FunctionComponent = () => {
  const [searchTerm, setSearchTerm] = React.useState<string>('');

  return (
    <>
      <SearchBox searchTermChanged={e => setSearchTerm(e.detail)}></SearchBox>
      <SearchResults
        entityTypes={['externalItem']}
        contentSources={['/external/connections/contosoProducts']}
        queryString={searchTerm}
      ></SearchResults>
    </>
  );
};
