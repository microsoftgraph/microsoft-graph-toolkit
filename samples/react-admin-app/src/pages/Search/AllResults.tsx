import { SearchBox, SearchResults } from '@microsoft/mgt-react/dist/es6/generated/react-preview';
import * as React from 'react';

const filesScopes = ['Files.Read.All', 'Sites.Read.All'];
const filesEntityTypes = ['driveItem', 'listItem', 'site'];
const bookmarksScopes = ['Bookmark.Read.All'];
const bookmarksEntityTypes = ['bookmark'];

export const AllResults: React.FunctionComponent = () => {
  const [searchTerm, setSearchTerm] = React.useState<string>('');

  return (
    <>
      <SearchBox searchTermChanged={e => setSearchTerm(e.detail)}></SearchBox>
      <SearchResults
        entityTypes={bookmarksEntityTypes}
        queryString={searchTerm}
        version="beta"
        size={1}
        scopes={bookmarksScopes}
      ></SearchResults>
      <SearchResults
        entityTypes={filesEntityTypes}
        queryString={searchTerm}
        scopes={filesScopes}
        fetchThumbnail={true}
      ></SearchResults>
    </>
  );
};
