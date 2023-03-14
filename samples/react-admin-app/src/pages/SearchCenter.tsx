import { IPivotStyles, IRawStyle, Label, Pivot, PivotItem } from '@fluentui/react';
import { IStyle } from '@fluentui/react-theme-provider';
import { SearchBox, SearchResults } from '@microsoft/mgt-react';
import * as React from 'react';
import { PageHeader } from '../components/PageHeader/PageHeader';
import { AllResults, ExternalItems, Interleaving, PeopleResults } from './Search';

export const SearchCenter: React.FunctionComponent = () => {
  return (
    <>
      <PageHeader
        title={'Search'}
        description={'Use this Search Center to test Microsot Graph Toolkit search components capabilities'}
      ></PageHeader>

      <Pivot styles={{ root: { paddingBottom: '10px' } }}>
        <PivotItem headerText="All Results">
          <AllResults></AllResults>
        </PivotItem>
        <PivotItem headerText="External Items">
          <ExternalItems></ExternalItems>
        </PivotItem>
        <PivotItem headerText="Interleaving">
          <Interleaving></Interleaving>
        </PivotItem>
        <PivotItem headerText="People">
          <PeopleResults></PeopleResults>
        </PivotItem>
      </Pivot>
    </>
  );
};
