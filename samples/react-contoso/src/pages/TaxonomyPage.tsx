import * as React from 'react';
import { PageHeader } from '../components/PageHeader';
import { TaxonomyExplorer } from './Taxonomy/TaxonomyExplorer';

const TaxonomyPage: React.FunctionComponent = () => {
  return (
    <>
      <PageHeader
        title={'Taxonomy Explorer'}
        description={'Use this taxonomy explorer to see all term groups, term sets and terms available'}
      ></PageHeader>
      <TaxonomyExplorer />
    </>
  );
};

export default TaxonomyPage;
