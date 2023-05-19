import { IResultTemplates } from './IResultTemplates';
import { IDataFilterResult } from './IDataFilter';

export interface IDataSourceData {
  /**
   * Items returned by the data source.
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  items: any[];

  /**
   * The count of items returned by the datasource
   */
  totalCount?: number;

  /**
   * The available filters provided by the data source according to the filters configuration provided from the data context (if applicable).
   */
  filters?: IDataFilterResult[];

  /**
   * Result templates available for items provided by the data source
   */
  resultTemplates?: IResultTemplates;
}
