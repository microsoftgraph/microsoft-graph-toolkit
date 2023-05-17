export class EventConstants {
  /**
   * Event name when filters are submitted
   */
  public static readonly SEARCH_FILTER_EVENT = 'Ubisoft:Components:Search:Filter';

  /**
   * Event name when results are retrieved
   */
  public static readonly SEARCH_RESULTS_EVENT = 'Ubisoft:Components:Search:Results';
  /**
   * Event name when input is sent
   */
  public static readonly SEARCH_INPUT_EVENT = 'Ubisoft:Components:Search:Input';

  /**
   * Event name when tab change is retrieved
   */
  public static readonly SEARCH_VERTICAL_EVENT = 'Ubisoft:Components:Search:Verticals';

  /**
   * Event name when sort field is updated
   */
  public static readonly SEARCH_SORT_EVENT = 'Ubisoft:Components:Search:Sort';

  /**
   * Event name when export profiles button is clicked
   */
  public static readonly SEARCH_EXPORT_PROFILES_EVENT = 'Ubisoft:Components:Search:ExportProfiles';
}

export class AnalyticsEventConstants {
  public static readonly MONITORED_EVENT = 'Ubisoft:Components:Search:MonitoredEvent';
}

export enum ComponentElements {
  UbisoftAuthProviderElement = 'ubisoft-auth-provider',
  UbisoftSearchResultsElement = 'ubisoft-search-results',
  UbisoftSearchFiltersElement = 'ubisoft-search-filters',
  UbisoftCheckboxFilterComponentElement = 'ubisoft-filter-checkbox',
  UbisoftDateFilterComponentElement = 'ubisoft-filter-date',
  UbisoftSearchSortComponentElement = 'ubisoft-search-sort',
  UbisoftProfileSocialComponentElement = 'ubisoft-profile-social',
  UbisoftPaginationElement = 'ubisoft-pagination',
  UbisoftSearchInputComponent = 'ubisoft-search-input',
  UbisoftSearchVerticalsComponent = 'ubisoft-search-verticals',
  UbisoftInputAutocompleteComponent = 'ubisoft-input-autocomplete',
  UbisoftAdaptiveCard = 'ubisoft-adaptive-card',
  UbisoftLanguageProvider = 'ubisoft-language-provider',
  UbisoftSearchInfos = 'ubisoft-search-infos',
  UbisoftErrorMEssage = 'ubisoft-error-message',
  UbisoftExportProfiles = 'ubisoft-export-profiles',
  UbisoftAnalytics = 'ubisoft-analytics'
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
