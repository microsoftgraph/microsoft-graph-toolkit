export class EventConstants {
  /**
   * Event name when filters are submitted
   */
  public static readonly SEARCH_FILTER_EVENT = 'Mgt:Components:Search:Filter';

  /**
   * Event name when results are retrieved
   */
  public static readonly SEARCH_RESULTS_EVENT = 'Mgt:Components:Search:Results';
  /**
   * Event name when input is sent
   */
  public static readonly SEARCH_INPUT_EVENT = 'Mgt:Components:Search:Input';

  /**
   * Event name when tab change is retrieved
   */
  public static readonly SEARCH_VERTICAL_EVENT = 'Mgt:Components:Search:Verticals';
  /**
   * Event name when sort field is updated
   */
  public static readonly SEARCH_SORT_EVENT = 'Mgt:Components:Search:Sort';
}

export enum ComponentElements {
  MgtAuthProviderElement = 'mgt-auth-provider',
  MgtSearchResultsElement = 'mgt-search-results',
  MgtSearchFiltersElement = 'mgt-search-filters',
  MgtCheckboxFilterComponentElement = 'mgt-filter-checkbox',
  MgtDateFilterComponentElement = 'mgt-filter-date',
  MgtSearchSortComponentElement = 'mgt-search-sort',
  MgtProfileSocialComponentElement = 'mgt-profile-social',
  MgtPaginationElement = 'mgt-pagination',
  MgtSearchInputComponent = 'mgt-search-input',
  MgtSearchVerticalsComponent = 'mgt-search-verticals',
  MgtInputAutocompleteComponent = 'mgt-input-autocomplete',
  MgtAdaptiveCard = 'mgt-adaptive-card',
  MgtLanguageProvider = 'mgt-language-provider',
  MgtSearchInfos = 'mgt-search-infos',
  MgtErrorMEssage = 'mgt-error-message',
  MgtExportProfiles = 'mgt-export-profiles',
  MgtAnalytics = 'mgt-analytics'
}

export enum ThemeCSSVariables {
  fontFamilyPrimary = '--ubi365-fontFamilyPrimary',
  fontFamilySecondary = '--ubi365-fontFamilySecondary',
  colorPrimary = '--ubi365-colorPrimary',
  colorSecondary = '--ubi365-colorSecondary',
  colorLight = '--ubi365-colorLight',
  textColor = '--text-color',
  //Grays
  light100 = '--gray100',
  light300 = '--gray300',
  //Tabs
  tabShadowHover = '--tab-border-color',
  borderTabs = '--tabs-border-colors',
  //Results
  topicBackground = '--topic-background-color',
  topicHover = '--topic-hover-color',
  topicFocus = '--topic-focus-color',
  topicDisable = '--topic-disable-color',
  //Recommended
  recommendedBorder = '--recommended-border',
  recommendedBackground = '--recommended-background'
}

export class CSSVariables {
  public static readonly ThemeVariables = [
    ThemeCSSVariables.fontFamilyPrimary,
    ThemeCSSVariables.fontFamilySecondary,
    ThemeCSSVariables.colorLight,
    ThemeCSSVariables.colorSecondary,
    ThemeCSSVariables.colorPrimary
  ];
}
