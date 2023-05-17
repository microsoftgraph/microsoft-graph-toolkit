import { BuiltinFilterTemplates } from './BuiltinTemplate';
import { FilterConditionOperator } from './IDataFilter';
import { ILocalizedString } from './ILocalizedString';

export enum FilterType {
  /**
   * A 'Refiner' filter means the filter gets the filters values from the data source and sends back the selected ones the data source.
   */
  Refiner = 'refiner',

  /**
   * An 'Static' filter means the filter doesn't care about filter values sent by the data source and provides its own arbitrary values regardless of input values.
   * A date picker or a taxonomy picker (or any picker) are good examples of what an 'Out' filter is.
   */
  StaticFilter = 'staticFilter'
}

export enum FilterSortType {
  /**
   * Sort filter values by their count
   */
  ByCount = 'byCount',
  /**
   * Sort filter values by their display name
   */
  ByName = 'byName'
}

export enum FilterSortDirection {
  Ascending = 'ascending',
  Descending = 'descending'
}

export interface IDataFilterConfiguration {
  /**
   * The internal filter name (ex: corresponding either to data source field name or refinable search managed property in the case of SharePoint)
   */
  filterName: string;

  /**
   * The flter name to display in the UI
   */
  displayName: string | ILocalizedString;

  /**
   * The template to use to show filters
   */
  template: BuiltinFilterTemplates;

  /**
   * Specifies if the filter should show values count
   */
  showCount: boolean;

  /**
   * The operator to use between filter values
   */
  operator: FilterConditionOperator;

  /**
   * Indicates if the filter allows multi values
   */
  isMulti: boolean;

  /**
   * If the filter should be sorted by name or by count
   */
  sortBy: FilterSortType;

  /**
   * The filter values sort direction (ascending/descending)
   */
  sortDirection: FilterSortDirection;

  /**
   * The index of this filter in the configuration
   */
  sortIdx: number;

  /**
   * Number of buckets to fetch
   */
  maxBuckets: number;

  /**
   * Aggregations to use for filter values
   */
  aggregations?: IDataFilterAggregation[];
}

export interface IDataFilterAggregation {
  /**
   * The friendly name to display as filter value for user (ex: "Word document")
   */
  aggregationName: string | ILocalizedString;

  /**
   * The values matching the aggreagation (ex: ["docx","doc"])
   */
  matchingValues: string[];

  /**
   * FQL value to apply when clicked: ex: or("value1","value2")
   */
  aggregationValue: string;

  /**
   * The icon URL to display for that value. Optional
   */
  aggregationValueIconUrl?: string;
}
