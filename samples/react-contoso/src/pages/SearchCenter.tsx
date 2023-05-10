import { IPivotStyleProps, IPivotStyles, IStyleFunctionOrObject, Pivot, PivotItem } from '@fluentui/react';
import * as React from 'react';
import { PageHeader } from '../components/PageHeader/PageHeader';
import { AllResults, ExternalItems, Interleaving, PeopleResults } from './Search';

export const SearchCenter: React.FunctionComponent = () => {
  const searchCenterStyles: React.CSSProperties = {
    maxWidth: '1028px',
    width: '100%',
    margin: 'auto'
  };

  const pivotStyles: IStyleFunctionOrObject<IPivotStyleProps, IPivotStyles> = {
    root: { paddingBottom: '10px' }
  };

  return (
    <>
      <PageHeader
        title={'Search'}
        description={'Use this Search Center to test Microsot Graph Toolkit search components capabilities'}
      ></PageHeader>

      <div style={searchCenterStyles}>
        <Pivot styles={pivotStyles}>
          <PivotItem headerText="All Results">
            <AllResults></AllResults>
          </PivotItem>
          <PivotItem headerText="Connectors">
            <ExternalItems></ExternalItems>
          </PivotItem>
          <PivotItem headerText="Interleaving">
            <Interleaving></Interleaving>
          </PivotItem>
          <PivotItem headerText="People">
            <PeopleResults></PeopleResults>
          </PivotItem>
        </Pivot>
      </div>
    </>
  );
};
