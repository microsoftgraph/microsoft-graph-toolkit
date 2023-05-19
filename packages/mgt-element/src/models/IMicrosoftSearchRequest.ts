export enum EntityType {
  Message = 'message',
  Event = 'event',
  Drive = 'drive',
  DriveItem = 'driveItem',
  ExternalItem = 'externalItem',
  List = 'list',
  ListItem = 'listItem',
  Site = 'site',
  Person = 'person',
  Bookmark = 'bookmark'
}

export interface IMicrosoftSearchQuery {
  requests: IMicrosoftSearchRequest[];
}

/**
 * https://docs.microsoft.com/en-us/graph/api/resources/searchrequest?view=graph-rest-beta
 */
export interface IMicrosoftSearchRequest {
  entityTypes: EntityType[];
  query: {
    queryString: string;
    queryTemplate?: string;
  };
  fields?: string[];
  aggregations?: ISearchRequestAggregation[];
  aggregationFilters?: string[];
  from?: number;
  size?: number;
  enableTopResults?: boolean;
  sortProperties?: ISearchSortProperty[];
  contentSources?: string[];
  queryAlterationOptions?: IQueryAlterationOptions;
  resultTemplateOptions?: {
    enableResultTemplate: boolean;
  };
  trimDuplicates?: boolean;
}

export interface ISearchSortProperty {
  name: string;
  isDescending: boolean;
}

export interface ISearchRequestAggregation {
  field: string;
  size?: number;
  bucketDefinition: IBucketDefinition;
}

export interface IBucketDefinition {
  sortBy: SearchAggregationSortBy;
  isDescending: boolean;
  minimumCount: number;
  ranges?: IBucketRangeDefinition[];
}

export enum SearchAggregationSortBy {
  Count = 'count',
  KeyAsNumber = 'keyAsNumber',
  KeyAsString = 'keyAsString'
}

export interface IBucketRangeDefinition {
  from?: number | string;
  to?: number | string;
}

export interface IQueryAlterationOptions {
  enableSuggestion: boolean;
  enableModification: boolean;
}
