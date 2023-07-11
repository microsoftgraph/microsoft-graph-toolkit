/* eslint-disable @typescript-eslint/class-literal-property-style */
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
  public static get SEARCH_INPUT_EVENT() {
    return 'Mgt:Components:Search:Input';
  }

  /**
   * Event name when tab change is retrieved
   */
  public static readonly SEARCH_VERTICAL_EVENT = 'Mgt:Components:Search:Verticals';
  /**
   * Event name when sort field is updated
   */
  public static readonly SEARCH_SORT_EVENT = 'Mgt:Components:Search:Sort';
}

export enum ThemeCSSVariables {
  fontFamilyPrimary = '--ubi365-fontFamilyPrimary',
  fontFamilySecondary = '--ubi365-fontFamilySecondary',
  colorPrimary = '--ubi365-colorPrimary',
  colorSecondary = '--ubi365-colorSecondary',
  colorLight = '--ubi365-colorLight'
}
