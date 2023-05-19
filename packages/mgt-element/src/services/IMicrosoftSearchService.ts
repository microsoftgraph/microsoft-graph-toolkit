import { IMicrosoftSearchDataSourceData } from '../models/IMicrosoftSearchDataSourceData';
import { IMicrosoftSearchQuery } from '../models/IMicrosoftSearchRequest';

export interface IMicrosoftSearchService {
  useBetaEndPoint: boolean;

  /**
   * Performs a search query against Microsoft Search
   * @param searchQuery The search query in KQL forma
   * @param culture The language cutlure to query (ex: 'fr-fr'). If not set, default is 'en-US')
   * @return The search results
   */
  search(searchQuery: IMicrosoftSearchQuery, culture?: string): Promise<IMicrosoftSearchDataSourceData>;

  /**
   * Abort the current HTTP request
   */
  abortRequest();
}
