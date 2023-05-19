import { FilterConditionOperator, IDataFilter } from '../models/IDataFilter';

export interface ISearchFiltersEventData {
  /**
   * List of filters to apply
   */
  selectedFilters: IDataFilter[];

  /**
   * Operator to use between filters
   */
  filterOperator: FilterConditionOperator;
}
