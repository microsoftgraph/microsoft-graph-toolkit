import { MgtTemplateProps, SearchBox, SearchResults } from '@microsoft/mgt-react';
import * as React from 'react';

export const ExternalItems: React.FunctionComponent = () => {
  const [searchTerm, setSearchTerm] = React.useState<string>('');

  return (
    <>
      <SearchBox searchTermChanged={e => setSearchTerm(e.detail)}></SearchBox>
      <SearchResults
        entityTypes={['externalItem']}
        contentSources={['/external/connections/contosopartsinventory']}
        queryString={searchTerm}
        scopes={['ExternalItem.Read.All']}
      >
        <ExternalItemTemplate template="result-externalItem"></ExternalItemTemplate>
      </SearchResults>
    </>
  );
};

const ExternalItemTemplate = (props: MgtTemplateProps) => {
  const { resource } = props.dataContext;
  return (
    <div>
      <h2>{resource.properties.name}</h2>
      {resource.properties.description}
    </div>
  );
};
