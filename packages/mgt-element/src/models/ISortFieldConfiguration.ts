import { ILocalizedString } from './ILocalizedString';

export enum SortFieldDirection {
  Ascending = 'asc',
  Descending = 'desc'
}

export interface ISortFieldConfiguration {
  /**
   * The sort search field to use
   */
  sortField: string;

  /**
   * The default sort direction
   */
  sortDirection: SortFieldDirection;

  /**
   * If the field is sorted by default (without use action)
   */
  isDefaultSort: boolean;

  /**
   * If the field should be available for users to sort
   */
  isUserSort: boolean;

  /**
   * If sortable by users, the display name to use in the control
   */
  sortFieldDisplayName: ILocalizedString;
}
