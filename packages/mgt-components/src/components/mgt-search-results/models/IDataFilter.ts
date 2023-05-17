/**
 * Represents a filter value to be send to a data source
 */
export interface IDataFilterValue {
  /**
   * An unique identifier for that filter value
   */
  key: string;

  /**
   * The filter value display name
   */
  name: string;

  /**
   * Inner value to use when the value is selected
   */
  value: string;

  /**
   * The comparison operator to use with this value. If not provided, the 'Equals' operator will be used.
   */
  operator?: FilterComparisonOperator;
}

/**
 * Represents a filter value returned by a data source
 */
export interface IDataFilterResultValue extends IDataFilterValue {
  /**
   * The number of results with this value
   */
  count: number;
}
/**
 * Represents a filter to be send to the data source
 */
export interface IDataFilter {
  /**
   * The filter internal name
   */
  filterName: string;

  /**
   * Values available in this filter
   */
  values: IDataFilterValue[];

  /**
   * The logical operator to use between values
   */
  operator?: FilterConditionOperator;
}

/**
 * Represents a filter returned from a data source
 */
export interface IDataFilterResult {
  /**
   * The filter display name
   */
  filterName: string;

  /**
   * Values available in this filter
   */
  values: IDataFilterResultValue[];
}
export enum FilterConditionOperator {
  OR = 'or',
  AND = 'and'
}
export enum FilterComparisonOperator {
  Eq = 0,
  Neq = 1,
  Gt = 2,
  Lt = 3,
  Geq = 4,
  Leq = 5,
  Contains = 6
}
