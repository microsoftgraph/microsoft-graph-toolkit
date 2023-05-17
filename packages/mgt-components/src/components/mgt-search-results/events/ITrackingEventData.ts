export interface ITrackingEventData {
  /**
   * The custom dimensions to submit with this event.
   */
  eventCustomDimensions?: IEventCustomDimension[]; // = Matomo custom dimension

  /**
   * The event category
   */
  category: EventCategory | string;

  /**
   * The event value
   */
  value?: number;

  /**
   * The associated user action
   */
  action: UserAction | string;

  /**
   * The event name
   */
  name: EventName | string;
}

export interface IEventCustomDimension {
  /**
   * Key too be resolved b y the proper Matomo ID for targeted application
   */
  key: SearchTrackedDimensions;

  /**
   * The data of that dimension
   */
  value: string;
}

export enum SearchTrackedDimensions {
  /* Action dimensions */
  SearchKeywords = 'SearchKeywords',
  ViewedResults = 'ViewedResults',
  SelectedVertical = 'SelectedVertical',
  SelectedFilter = 'SelectedFilter',
  ViewedBookmarks = 'ViewedBookmarks',
  ViewedPageNumbers = 'ViewedPageNumbers',
  SearchDataSources = 'SearchDataSources',
  SearchedContentType = 'SearchedContentType',
  ClickedItemRanks = 'ClickedItemRanks',

  /* Visit dimensions */
  UserJobFamily = 'UserJobFamily',
  UserLocation = 'Location',
  UserDepartment = 'Department'
}

export enum EventCategory {
  SearchResultsEvents = 'SearchResultsEvents',
  SearchFiltersEvents = 'SearchFiltersEvents',
  SearchBoxEvents = 'SearchBoxEvents',
  SearchVerticalsEvents = 'SearchVerticalsEvents'
}

export enum UserAction {
  SearchResultItemClicked = 'SearchResultItemClicked',
  SearchSuggestionClicked = 'SearchSuggestionClicked',
  SearchResultPageBrowsed = 'SearchResultPageBrowsed',
  SearchBookmarkClicked = 'SearchBookmarkClicked',
  SearchVerticalSelected = 'SearchVerticalSelected',
  SearchKeywordsSubmitted = 'SearchKeywordsSubmitted',
  SearchFilterValueSelected = 'SearchFilterValueSelected',
  SearchFilterValueReset = 'SearchFilterValueReset',
  SearchFilterResetAllValues = 'SearchFilterResetAllValues',
  SearchResultsDisplayed = 'SearchResultsDisplayed',
  SearchResultsNoResult = 'SearchResultsNoResult',
  SearchBookmarksDisplayed = 'SearchBookmarksDisplayed',
  SearchSuggestionsDisplayed = 'SearchSuggestionsDisplayed',
  SearchSortValueSelected = 'SearchSortValueSelected'
}

export enum EventName {
  InitialSearchKeywords = 'InitialSearchKeywords',
  NewSearchKeywords = 'NewSearchKeywords',
  InitialSearchVertical = 'InitialSearchVertical',
  NewSearchVertical = 'NewSearchVertical',
  NewSearchSort = 'NewSearchSort'
}
