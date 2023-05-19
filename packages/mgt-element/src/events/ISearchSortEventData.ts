import { ISearchSortProperty } from '../models/IMicrosoftSearchRequest';

export interface ISearchSortEventData {
  /**
   * The Microsoft Search properties
   */
  sortProperties: ISearchSortProperty[];
}
