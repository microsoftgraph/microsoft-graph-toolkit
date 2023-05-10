import { SearchBox, SearchResults } from '@microsoft/mgt-react';
import * as React from 'react';

export const AllResults: React.FunctionComponent = () => {
  const [searchTerm, setSearchTerm] = React.useState<string>('');

  return (
    <>
      <SearchBox searchTermChanged={e => setSearchTerm(e.detail)}></SearchBox>
      {searchTerm && (
        <>
          <SearchResults
            entityTypes={['bookmark']}
            queryString={searchTerm}
            version="beta"
            size={1}
            scopes={['Bookmark.Read.All']}
          ></SearchResults>
          <SearchResults
            entityTypes={['driveItem', 'listItem', 'site']}
            queryString={searchTerm}
            scopes={['Files.Read.All', 'Sites.Read.All']}
            fetchThumbnail={true}
          ></SearchResults>
        </>
      )}
    </>
  );
};
