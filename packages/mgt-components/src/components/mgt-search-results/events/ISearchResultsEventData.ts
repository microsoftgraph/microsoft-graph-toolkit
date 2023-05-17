import { IDataFilterResult } from '../models/IDataFilter';
import { IQueryAlterationResponse } from '../models/IMicrosoftSearchResponse';
import { ISortFieldConfiguration } from '../models/ISortFieldConfiguration';

export interface ISearchResultsEventData {
  availableFilters?: IDataFilterResult[];
  sortFieldsConfiguration?: ISortFieldConfiguration[];
  submittedQueryText?: string;
  resultsCount?: number;
  queryAlterationResponse?: IQueryAlterationResponse;
  from?: number;
}
