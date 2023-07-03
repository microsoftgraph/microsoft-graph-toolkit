import { IDataSourceData } from './IDataSourceData';
import { IQueryAlterationResponse } from './IMicrosoftSearchResponse';

export interface IMicrosoftSearchDataSourceData extends IDataSourceData {
  /**
   * The alteration response by Microsft Search
   * see https://docs.microsoft.com/en-us/graph/api/resources/alterationResponse?view=graph-rest-beta&preserve-view=true
   * Spelling correction is only supported for the following resources: message, event, site, drive, driveItem, list, listItem and externalItem.
   */
  queryAlterationResponse?: IQueryAlterationResponse;
}
