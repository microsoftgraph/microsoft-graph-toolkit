import { IMicrosoftSearchService } from './IMicrosoftSearchService';
import { IMicrosoftSearchDataSourceData } from '../models/IMicrosoftSearchDataSourceData';
import { EntityType, IMicrosoftSearchQuery } from '../models/IMicrosoftSearchRequest';
import { IMicrosoftSearchResponse, IMicrosoftSearchResultSet } from '../models/IMicrosoftSearchResponse';
import { IDataFilterResult, IDataFilterResultValue, FilterComparisonOperator } from '../models/IDataFilter';
import { Client, ClientOptions, FetchOptions } from '@microsoft/microsoft-graph-client';
import { Providers } from '../providers/Providers';

export const EntityTypesValidCombination = [
  EntityType.Drive,
  EntityType.DriveItem,
  EntityType.Site,
  EntityType.List,
  EntityType.ListItem
];

export class MicrosoftSearchService implements IMicrosoftSearchService {
  private _useBetaEndPoint: boolean;
  public get useBetaEndPoint(): boolean {
    return this._useBetaEndPoint;
  }
  public set useBetaEndPoint(value: boolean) {
    this._useBetaEndPoint = value;
    this._microsoftSearchUrl = `https://graph.microsoft.com/${value ? 'beta' : 'v1.0'}/search/query`;
  }

  private controller: AbortController;

  private _microsoftSearchUrl = 'https://graph.microsoft.com/v1.0/search/query';

  public async search(searchQuery: IMicrosoftSearchQuery, culture?: string): Promise<IMicrosoftSearchDataSourceData> {
    let itemsCount = 0;
    const response: IMicrosoftSearchDataSourceData = {
      items: [],
      filters: []
    };
    const aggregationResults: IDataFilterResult[] = [];

    // To be able to cancel requests individually, a new controller needs to be instanciated every time
    this.controller = new AbortController();

    const fetchOptions: FetchOptions = {
      signal: this.controller.signal
    };

    const clientOptions: ClientOptions = {
      authProvider: Providers.globalProvider,
      fetchOptions: fetchOptions
    };

    const client = Client.initWithMiddleware(clientOptions);

    try {
      const jsonResponse: IMicrosoftSearchResponse = await client
        .api(this._microsoftSearchUrl)
        .headers({
          SdkVersion: 'Ubisoft',
          'accept-language': culture ? culture : 'en-US'
        })
        .post(searchQuery);

      if (jsonResponse.value && Array.isArray(jsonResponse.value)) {
        jsonResponse.value.forEach((value: IMicrosoftSearchResultSet) => {
          // Map results
          value.hitsContainers.forEach(hitContainer => {
            itemsCount += hitContainer.total;

            if (hitContainer.hits) {
              const hits = hitContainer.hits.map(hit => {
                // "externalItem" will contain resource.properties but "listItem" will be resource.fields
                const propertiesFieldName = hit.resource.properties
                  ? 'properties'
                  : hit.resource.fields
                  ? 'fields'
                  : null;

                if (propertiesFieldName) {
                  // Flatten "fields" to be usable with the Search Fitler WP as refiners
                  Object.keys(hit.resource[propertiesFieldName]).forEach(field => {
                    // If the property already exists, keep it. Otherwise create it.
                    if (!hit[field.toLocaleLowerCase()]) {
                      hit[field.toLocaleLowerCase()] = hit.resource[propertiesFieldName][field];
                    }
                  });
                }

                return hit;
              });

              response.items = response.items.concat(hits);
            }

            if (hitContainer.aggregations) {
              // Map refinement results
              hitContainer.aggregations.forEach(aggregation => {
                const values: IDataFilterResultValue[] = [];
                aggregation.buckets.forEach(bucket => {
                  values.push({
                    key: bucket.aggregationFilterToken,
                    count: bucket.count,
                    name: bucket.key,
                    value: bucket.aggregationFilterToken,
                    operator: FilterComparisonOperator.Contains
                  } as IDataFilterResultValue);
                });

                aggregationResults.push({
                  filterName: aggregation.field,
                  values: values
                });
              });

              response.filters = aggregationResults;
            }
          });

          if (value?.queryAlterationResponse) {
            response.queryAlterationResponse = value.queryAlterationResponse;
          }

          if (value?.resultTemplates) {
            response.resultTemplates = value.resultTemplates;
          }
        });

        response.totalCount = itemsCount;
      }
    } catch (error) {
      if (this.controller.signal.aborted) {
        console.log('Graph request was aborted');
      }

      throw new Error(`${error.code} - ${error.message}`);
    }

    return response;
  }

  public abortRequest() {
    this.controller.abort();
  }
}
