import { LocalizationHelper, Providers, ProviderState, MgtConnectableComponent } from '@microsoft/mgt-element';
import { css, customElement, html, property, PropertyValues, state } from 'lit-element';
import { IDataSourceData } from './models/IDataSourceData';
import {
  IMicrosoftSearchQuery,
  SearchAggregationSortBy,
  ISearchRequestAggregation,
  EntityType,
  ISearchSortProperty
} from './models/IMicrosoftSearchRequest';
import { IMicrosoftSearchService } from './services/IMicrosoftSearchService';
import { ITemplateService } from './services/ITemplateService';
import { MicrosoftSearchService } from './services/MicrosoftSearchService';
import { TemplateService } from './services/TemplateService';
import { isEmpty } from 'lodash-es';
import { ISearchFiltersEventData } from './events/ISearchFiltersEventData.ts';
import { DataFilterHelper } from './helpers/DataFilterHelper';
import { ISearchResultsEventData } from './events/ISearchResultsEventData';
import { AnalyticsEventConstants, ComponentElements, EventConstants } from './Constants';
import { FilterSortDirection, FilterSortType } from './models/IDataFilterConfiguration';
//import { MgtSearchFiltersComponent } from "../Mgt-search-filters/Mgt-search-filters";
import { BuiltinFilterTemplates } from './models/BuiltinTemplate';
import { DateHelper } from './helpers/DateHelper';
//import { MgtPaginationComponent } from "../../internal/Mgt-pagination/Mgt-pagination";
//import { MgtSearchInputComponent } from "../Mgt-search-input/Mgt-search-input";
import { unsafeHTML } from 'lit-html/directives/unsafe-html';
import { SearchResultsHelper } from './helpers/SearchResultsHelper';
import { ISearchVerticalEventData } from './events/ISearchVerticalEventData';
//import { MgtSearchVerticalsComponent } from "../Mgt-search-verticals/Mgt-search-verticals";
import { repeat } from 'lit-html/directives/repeat';
import { ISearchInputEventData } from './events/ISearchInputEventData';
import { UrlHelper } from './helpers/UrlHelper';
import { MgtSearchResultsStrings as strings } from './loc/strings.default';
import { ILocalizedString } from './models/ILocalizedString';
import { nothing } from 'lit-html';
import { BuiltinTokenNames, TokenService } from './services/TokenService';
import { ITokenService } from './services/ITokenService';
import { UserAction, EventCategory, ITrackingEventData, SearchTrackedDimensions } from './events/ITrackingEventData';
import { IMicrosoftSearchDataSourceData } from './models/IMicrosoftSearchDataSourceData';
//import { MgtErrorMessageComponent } from "../../../components/internal/Mgt-error-message/Mgt-error-message";
import { ISortFieldConfiguration, SortFieldDirection } from './models/ISortFieldConfiguration';
import { ISearchSortEventData } from './events/ISearchSortEventData';
import { styles as tailwindStyles } from '../../styles/tailwind-styles-css';

@customElement('mgt-search-results')
export class MgtSearchResultsComponent extends MgtConnectableComponent {
  //#region Attributes

  /**
   * Flag indicating if the beta endpoint for Microsoft Graph API should be used
   */
  @property({ type: Boolean, attribute: 'use-beta' })
  useBetaEndpoint = false;

  /**
   * The Microsoft Search entity types to query
   */
  @property({
    type: String,
    attribute: 'entity-types',
    converter: {
      fromAttribute: value => {
        return value.split(',') as EntityType[];
      }
    }
  })
  entityTypes: EntityType[] = [EntityType.ListItem];

  /**
   * The default query text to apply.
   * Query string parameter and search box have priority over this value during first load
   */
  @property({ type: String, attribute: 'query-text' })
  defaultQueryText: string;

  /**
   * The search query template to use. Support tokens https://learn.microsoft.com/en-us/graph/search-concept-query-template
   */
  @property({ type: String, attribute: 'query-template' })
  queryTemplate: string;

  /**
   * If specified, get the default query text from this query string parameter name
   */
  @property({ type: String, attribute: 'default-query-string-parameter' })
  defaultQueryStringParameter: string;

  /**
   * Search managed properties to retrieve for results and usable in the results template.
   * Comma separated. Refer to the [Microsoft Search API documentation](https://learn.microsoft.com/en-us/graph/api/resources/search-api-overview?view=graph-rest-1.0&preserve-view=true#scope-search-based-on-entity-types) to know what properties can be used according to entity types.
   */
  @property({
    type: Array,
    attribute: 'fields',
    converter: {
      fromAttribute: value => {
        return value.split(',');
      }
    }
  })
  selectedFields: string[] = [
    'name',
    'title',
    'summary',
    'created',
    'createdBy',
    'filetype',
    'defaultEncodingURL',
    'lastModifiedTime',
    'modifiedBy',
    'path',
    'hitHighlightedSummary',
    'SPSiteURL',
    'SiteTitle'
  ];

  /**
   * Sort properties for the request
   */
  @property({
    type: String,
    attribute: 'sort-properties',
    converter: {
      fromAttribute: value => {
        try {
          return JSON.parse(value);
        } catch {
          return null;
        }
      }
    }
  })
  sortFieldsConfiguration: ISortFieldConfiguration[] = [];

  /**
   * Flag indicating if the pagniation control should be displayed
   */
  @property({ type: Boolean, attribute: 'show-paging' })
  showPaging: boolean;

  /**
   * The number of results to show per results page
   */
  @property({ type: Number, attribute: 'page-size' })
  pageSize = 10;

  /**
   * The number of pages to display in the pagination control
   */
  @property({ type: Number, attribute: 'pages-number' })
  numberOfPagesToDisplay = 5;

  /**
   * Flag indicating if Micrsoft Search result types should be applied in results
   */
  @property({ type: Boolean, attribute: 'enable-result-types' })
  enableResultTypes: boolean;

  /**
   * If "entityTypes" contains "externalItem", specify the connection id of the external source
   */
  @property({
    type: String,
    attribute: 'connections',
    converter: {
      fromAttribute: value => {
        return value.split(',').map(v => `/external/connections/${v}`) as string[];
      }
    }
  })
  connectionIds: string[];

  /**
   * Indicates whether spelling modifications are enabled. If enabled, the user will get the search results for the corrected query in case of no results for the original query with typos.
   */
  @property({ type: Boolean, attribute: 'enable-modification' })
  enableModification = false;

  /**
   * Indicates whether spelling suggestions are enabled. If enabled, the user will get the search results for the original search query and suggestions for spelling correction
   */
  @property({ type: Boolean, attribute: 'enable-suggestion' })
  enableSuggestion = false;

  /**
   * If specified, shows the title on top of the results
   */
  @property({
    type: String,
    attribute: 'comp-title',
    converter: {
      fromAttribute: value => {
        try {
          return JSON.parse(value);
        } catch {
          return value;
        }
      }
    }
  })
  componentTitle: string | ILocalizedString;

  /**
   * If specified, shows a "See all" link at top top of the results
   */
  @property({ type: String, attribute: 'see-all-link' })
  seeAllLink: string;

  /**
   * If specified, show the results count at the top of the results
   */
  @property({ type: Boolean, attribute: 'show-count' })
  showCount: boolean;

  /**
   * The search filters component ID if connected to a search filters
   */
  @property({ type: String, attribute: 'search-filters-id' })
  searchFiltersComponentId: string;

  /**
   * The search input component ID if connected to a search input
   */
  @property({ type: String, attribute: 'search-input-id' })
  searchInputComponentId: string;

  /**
   * The search verticals component ID if connected to a search verticals
   */
  @property({ type: String, attribute: 'search-verticals-id' })
  searchVerticalsComponentId: string;

  /**
   * The search sort component ID if connected to a search sort component
   */
  @property({ type: String, attribute: 'search-sort-id' })
  searchSortComponentId: string;

  /**
   * If connected to a search verticals component on the same page, determines on which keys this component should be displayed
   */
  @property({
    type: Array,
    attribute: 'verticals-keys',
    converter: {
      fromAttribute: value => {
        return value.split(',');
      }
    }
  })
  selectedVerticalKeys: string[];

  /**
   * Flag indicating if the loading indication (spinner/shimmers) should be displayed when fectching the data
   */
  @property({ type: Boolean, attribute: 'no-loading' })
  noLoadingIndicator: boolean;

  //#endregion

  //#region State properties

  @state()
  data: IDataSourceData = { items: [] };

  @state()
  private isLoading = true;

  @state()
  private shouldRender: boolean;

  @state()
  private error: Error = null;

  //#endregion

  //#region Class properties
  public declare searchQuery: IMicrosoftSearchQuery;
  public declare msSearchService: IMicrosoftSearchService;
  private declare templateService: ITemplateService;
  private declare tokenService: ITokenService;
  private declare dateHelper: DateHelper;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private declare dayJs: any;
  private declare currentLanguage: string;
  declare sortProperties: ISearchSortProperty[];
  //#endregion

  //#region MGT/Lit Lifecycle methods

  constructor() {
    super();
    this.msSearchService = new MicrosoftSearchService();
    this.templateService = new TemplateService();
    this.tokenService = new TokenService();

    this.dateHelper = new DateHelper(LocalizationHelper.strings?.language);

    this.searchQuery = {
      requests: []
    };

    this.addEventListener('templateRendered', (e: CustomEvent) => {
      const element = e.detail.element as HTMLElement;

      if (this.enableResultTypes) {
        // Process result types and replace part of HTML with item id
        /*const newElement = this.templateService.processResultTypesFromHtml(this.data, element, this.getTheme());
                element.replaceWith(newElement);*/
      }
    });

    this.handleSearchVertical = this.handleSearchVertical.bind(this);
    this.handleSearchFilters = this.handleSearchFilters.bind(this);
    this.handleSearchInput = this.handleSearchInput.bind(this);
    this.handleSearchSort = this.handleSearchSort.bind(this);

    this.goToPage = this.goToPage.bind(this);
  }

  public render() {
    if (this.shouldRender) {
      let renderHeader;
      let renderItems;
      let renderOverlay;
      let renderPagination;

      // Render shimmers
      if (this.hasTemplate('shimmers') && !this.noLoadingIndicator && !this.renderedOnce) {
        renderItems = this.renderTemplate('shimmers', { items: Array(this.pageSize) });
      } else {
        // Render loading overlay
        if (this.isLoading && !this.noLoadingIndicator) {
          renderOverlay = html`
                        <div class="absolute bg-white bg-opacity-60 h-full w-full flex items-center justify-center">
                            <div class="flex items-center relative">
                                <svg class="animate-spin h-8 w-8 text-primary" xmlns="http://www.w3.org/2000/svg" fill="none"
                                    viewBox="0 0 24 24">
                                    <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle>
                                    <path class="opacity-75" fill="currentColor"
                                    d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z">
                                    </path>
                                </svg>
                            </div>
                        </div>
                    `;
        }
      }

      // Render error
      if (this.error) {
        return html`<Mgt-error-message .error=${this.error}></Mgt-error-message>`;
      }

      // Render header
      if (this.componentTitle || this.seeAllLink || this.showCount) {
        renderHeader = html`
                    <div class="font-Mgt flex items-end justify-between mb-4">
                        <div class="space-x-2">
                        ${
                          this.componentTitle
                            ? html`
                                <div data-ref="component-title" class="text-3xl inline-block selection:tracking-[0.0012em] font-bold text-transparent bg-clip-text bg-gradient-to-r from-gradientFrom to-gradientTo">${this.getLocalizedString(
                                  this.componentTitle
                                )}</div>
                            `
                            : null
                        }
                        ${
                          this.showCount && !this.isLoading
                            ? html`
                                <div data-ref="show-count" class="text-sm inline-block font-normal font-sans">${this.data.totalCount} ${strings.results}</div>
                            `
                            : null
                        }                            
                        </div>
                        ${
                          this.seeAllLink && this.data.totalCount > 0
                            ? html`
                            <div class="text-sm text-primary">
                                <a data-ref="see-all-link" class="flex items-center rounded hover:text-primaryHover focus:outline focus:outline-2 focus-visible:outline focus-visible:outline-2" href="${this.tokenService.resolveTokens(
                                  this.seeAllLink
                                )}" title=${strings.seeAllLink}>
                                    <span>${strings.seeAllLink}</span>
                                    <svg width="18" height="15" viewBox="0 0 18 18" class="fill-current ml-4" xmlns="http://www.w3.org/2000/svg">
                                        <path d="M11.526 14.29C11.526 14.29 16.027 9.785 17.781 8.03C17.927 7.884 18 7.69199 18 7.49999C18 7.30799 17.927 7.11699 17.781 6.97C16.028 5.21599 11.526 0.711995 11.526 0.711995C11.382 0.566995 11.192 0.494995 11.002 0.494995C10.809 0.494995 10.617 0.568995 10.47 0.715995C10.177 1.00799 10.175 1.482 10.466 1.772L15.444 6.74999H0.751953C0.337953 6.74999 0.00195312 7.08599 0.00195312 7.49999C0.00195312 7.91399 0.337953 8.24999 0.751953 8.24999H15.444L10.465 13.229C10.176 13.518 10.179 13.991 10.471 14.283C10.619 14.431 10.812 14.505 11.004 14.505C11.194 14.505 11.382 14.433 11.526 14.29Z"/>
                                    </svg>
                                </a>
                            </div>
                        `
                            : null
                        }
                    </div>
                `;
      }

      if (this.renderedOnce) {
        // Render items
        if (this.hasTemplate('items')) {
          renderItems = this.renderTemplate('items', this.data);
        } else {
          // Default template for all items
          renderItems = html`
                        <ul class="space-y-4">
                            ${repeat(
                              this.data.items,
                              item => item.hitId,
                              item => {
                                return html`             
                                        <li id=${item.hitId} data-ref="item" class="!mt-0 !mb-8 p-2">
                                            <div class="flex items-center space-x-2 text-2xl mb-2">
                                                <egg-icon icon-id="egg-global:file:doc"></egg-icon>
                                                <a href=${
                                                  item?.resource?.fields?.defaultEncodingURL
                                                } class="hover:text-primary">
                                                    <span class="font-Mgt font-bold">${SearchResultsHelper.getItemTitle(
                                                      item
                                                    )}</span>
                                                </a>
                                            </div>
                                            <div class="font-sans text-sm text-black/[0.6] mb-2">
                                                <span class="itemInfo">${this.dayJs(item.resource?.created).format(
                                                  'DD/MM/YYYY'
                                                )}</span>
                                                <span class="itemInfo">By <a href="" class="font-bold hover:text-primary">${
                                                  item.resource?.createdBy?.user?.displayName
                                                }</a></span>
                                                <a href="" class="bg-topicBackground py-1 px-4 rounded-[13px] text-xs font-Mgt text-textColor transition-all cursor-pointer hover:bg-topicHover focus:bg-topicFocus focus-visible:bg-topicFocus">${
                                                  item.resource?.fields?.siteTitle
                                                }</a>
                                            </div>
                                            <div class="font-sans">
                                                <p class="itemInfo text-base color-textColor line-clamp-2 overflow-hidden" style="display: -webkit-box;-webkit-box-orient: vertical;-webkit-line-clamp: 2;">${unsafeHTML(
                                                  SearchResultsHelper.getItemSummary(item.summary)
                                                )}</p>
                                            </div>
                                        </li>
                                    `;
                              }
                            )}
                        </ul>
                    `;
        }

        // Render pagination
        if (this.showPaging && this.data.items.length > 0) {
          renderPagination = html`<Mgt-pagination 
                                                class="flex justify-center p-2 mt-6 mb-6 min-w-full" 
                                                .totalItems=${this.data.totalCount} 
                                                .itemsCountPerPage=${this.pageSize} 
                                                .numberOfPagesToDisplay=${this.numberOfPagesToDisplay}
                                                .onPageNumberUpdated=${this.goToPage}
                                            >
                                            </Mgt-pagination>
                                        `;
        }
      }

      return html`
                ${renderHeader}            
                <div class="relative flex justify-between flex-col">
                    ${renderOverlay}
                    ${renderItems}
                    ${renderPagination}
                </div>         
            `;
    }

    return nothing;
  }

  public async connectedCallback(): Promise<void> {
    // 'setTimeout' is used here to make sure the initialization logic for the search results components occurs after other component initialization routine.
    // This way, we ensure other component properties will be accessible.
    // connectedCallback events on other components will execute before according to the JS event loop
    // https://developer.mozilla.org/en-US/docs/Web/JavaScript/EventLoop
    setTimeout(async () => {
      this.msSearchService.useBetaEndPoint = this.useBetaEndpoint;
      this.dayJs = await this.dateHelper.dayJs();

      // Bind connected components
      const bindings = [
        {
          id: this.searchVerticalsComponentId,
          eventName: EventConstants.SEARCH_VERTICAL_EVENT,
          callbackFunction: this.handleSearchVertical
        },
        {
          id: this.searchFiltersComponentId,
          eventName: EventConstants.SEARCH_FILTER_EVENT,
          callbackFunction: this.handleSearchFilters
        },
        {
          id: this.searchInputComponentId,
          eventName: EventConstants.SEARCH_INPUT_EVENT,
          callbackFunction: this.handleSearchInput
        },
        {
          id: this.searchSortComponentId,
          eventName: EventConstants.SEARCH_SORT_EVENT,
          callbackFunction: this.handleSearchSort
        }
      ];

      this.bindComponents(bindings);

      if (this.enableResultTypes) {
        // Only load adaptive cards bundle if result types are enabled for performance purpose
        await this.templateService.loadAdaptiveCardsResources();
      }

      if (this.searchVerticalsComponentId) {
        // Check if the current component should be displayed at first
        const verticalsComponent = document.getElementById(this.searchVerticalsComponentId) as any; // MgtSearchVerticalsComponent;
        if (verticalsComponent) {
          // Reead the default value directly from the attribute
          const selectedVerticalKey = verticalsComponent.selectedVerticalKey;
          if (selectedVerticalKey) {
            this.shouldRender = this.selectedVerticalKeys.indexOf(selectedVerticalKey) !== -1;
          }
        }
      } else {
        this.shouldRender = true;
      }

      // Set default sort properties according to configuration
      this.initSortProperties();

      // Build the search query
      this.buildSearchQuery();

      // Set tokens
      this.tokenService.setTokenValue(BuiltinTokenNames.searchTerms, this.getDefaultQueryText());

      return super.connectedCallback();
    });
  }

  public disconnectedCallback(): void {
    super.disconnectedCallback();
  }

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  public updated(changedProperties: PropertyValues<this>): void {
    // Process result types on default Lit template of this component
    if (this.enableResultTypes && !this.hasTemplate('items')) {
      this.templateService.processResultTypesFromHtml(this.data, this.renderRoot as HTMLElement);
    }

    // Properties trigerring a new search
    // Mainly use for Storybook demo scenario
    if (
      changedProperties.get('defaultQueryText') ||
      changedProperties.get('selectedFields') ||
      changedProperties.get('pageSize') ||
      changedProperties.get('entityTypes') ||
      changedProperties.get('enableResultTypes') ||
      changedProperties.get('connectionIds') ||
      changedProperties.get('numberOfPagesToDisplay')
    ) {
      // Update the search query
      this.buildSearchQuery();
      this._search(this.searchQuery);
    }

    this.currentLanguage = LocalizationHelper.strings?.language;
  }

  /**
   * Only calls when the provider is in ProviderState.SignedIn state
   * @returns
   */
  public async loadState(): Promise<void> {
    if (this.shouldRender && this.getDefaultQueryText()) {
      await this._search(this.searchQuery);
    } else {
      this.isLoading = false;
    }
  }

  //#endregion

  //#region Static properties accessors
  static get styles() {
    return [
      css`
      :host  {
  
          .itemInfo + .itemInfo::before {
              content: " • ";
              font-size: 18px;
              line-height: 1;
              transform: translateY(2px);
              display: inline-block;
              margin: 0 5px;
          }

          .itemInfo + .itemInfo::after {
              content: " • ";
              font-size: 18px;
              line-height: 1;
              transform: translateY(2px);
              display: inline-block;
              margin: 0 5px;
          }
      }                 
      `,
      tailwindStyles
    ];
  }

  protected get strings() {
    return strings;
  }

  //#endregion

  //#region Data related methods

  private async _search(searchQuery: IMicrosoftSearchQuery): Promise<void> {
    const provider = Providers.globalProvider;
    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    try {
      // Reset error
      this.error = null;

      this.isLoading = true;
      const queryLanguage = LocalizationHelper.strings?.language;

      const results = await this.msSearchService.search(searchQuery, queryLanguage);
      this.data = results;

      // Enhance results
      this.data.items = SearchResultsHelper.enhanceResults(this.data.items);
      this.isLoading = false;

      this.renderedOnce = true;

      // Notify subscribers new filters are available
      this.fireCustomEvent(EventConstants.SEARCH_RESULTS_EVENT, {
        availableFilters: results.filters,
        sortFieldsConfiguration: this.sortFieldsConfiguration.filter(s => s.isUserSort),
        submittedQueryText: this.searchQuery.requests[0].query.queryString,
        resultsCount: results.totalCount,
        queryAlterationResponse: results.queryAlterationResponse,
        from: this.searchQuery.requests[0].from
      } as ISearchResultsEventData);

      // Track events for analytics
      this.trackEvents(results);
    } catch (error) {
      this.error = error;
    }
  }

  private buildAggregationsFromFiltersConfig(): ISearchRequestAggregation[] {
    let aggregations: ISearchRequestAggregation[] = [];
    const filterComponent = document.getElementById(this.searchFiltersComponentId) as any; //MgtSearchFiltersComponent;
    if (filterComponent && filterComponent.filterConfiguration) {
      // Build aggregations from filters configuration (i.e. refiners)
      aggregations = filterComponent.filterConfiguration.map(filterConfiguration => {
        const aggregation: ISearchRequestAggregation = {
          field: filterConfiguration.filterName,
          bucketDefinition: {
            isDescending: filterConfiguration.sortDirection === FilterSortDirection.Ascending ? false : true,
            minimumCount: 0,
            sortBy:
              filterConfiguration.sortBy === FilterSortType.ByCount
                ? SearchAggregationSortBy.Count
                : SearchAggregationSortBy.KeyAsString
          },
          size: filterConfiguration && filterConfiguration.maxBuckets ? filterConfiguration.maxBuckets : 10
        };

        if (filterConfiguration.template === BuiltinFilterTemplates.Date) {
          const pastYear = this.dayJs(new Date()).subtract(1, 'years').subtract(1, 'minutes').toISOString();
          const past3Months = this.dayJs(new Date()).subtract(3, 'months').subtract(1, 'minutes').toISOString();
          const pastMonth = this.dayJs(new Date()).subtract(1, 'months').subtract(1, 'minutes').toISOString();
          const pastWeek = this.dayJs(new Date()).subtract(1, 'week').subtract(1, 'minutes').toISOString();
          const past24hours = this.dayJs(new Date()).subtract(24, 'hours').subtract(1, 'minutes').toISOString();
          const today = new Date().toISOString();

          aggregation.bucketDefinition.ranges = [
            {
              to: pastYear
            },
            {
              from: pastYear,
              to: today
            },
            {
              from: past3Months,
              to: today
            },
            {
              from: pastMonth,
              to: today
            },
            {
              from: pastWeek,
              to: today
            },
            {
              from: past24hours,
              to: today
            },
            {
              from: today
            }
          ];
        }

        return aggregation;
      });
    }

    return aggregations;
  }

  /**
   * Builds the search query according to the current componetn parameters and context
   */
  private buildSearchQuery() {
    // Build base search query from parameters
    this.searchQuery = {
      requests: [
        {
          entityTypes: this.entityTypes,
          contentSources: this.connectionIds,
          fields: this.selectedFields,
          query: {
            queryString: this.getDefaultQueryText(),
            queryTemplate: this.queryTemplate
          },
          from: 0,
          size: this.pageSize,
          queryAlterationOptions: {
            enableModification: this.enableModification,
            enableSuggestion: this.enableSuggestion
          },
          resultTemplateOptions: {
            enableResultTemplate: this.enableResultTypes
          }
        }
      ]
    };

    // Sort properties
    if (this.sortProperties && this.sortProperties.length > 0) {
      this.searchQuery.requests[0].sortProperties = this.sortProperties;
    }

    // If a filter component is connected, get the configuration directly from connected component
    if (this.searchFiltersComponentId) {
      this.searchQuery.requests[0].aggregations = this.buildAggregationsFromFiltersConfig();
    }
  }

  //#endregion

  //#region Event handlers from connected components

  private async handleSearchFilters(e: CustomEvent<ISearchFiltersEventData>): Promise<void> {
    if (this.shouldRender) {
      let aggregationFilters: string[] = [];

      const selectedFilters = e.detail.selectedFilters;

      // Build aggregation filters
      if (selectedFilters.some(f => f.values.length > 0)) {
        // Bind to current context to be able to refernce "dayJs"
        const buildFqlRefinementString = DataFilterHelper.buildFqlRefinementString.bind(this);

        // Make sure, if we have multiple filters, at least two filters have values to avoid apply an operator ("or","and") on only one condition failing the query.
        if (
          selectedFilters.length > 1 &&
          selectedFilters.filter(selectedFilter => selectedFilter.values.length > 0).length > 1
        ) {
          const refinementString = buildFqlRefinementString(selectedFilters).join(',');
          if (!isEmpty(refinementString)) {
            aggregationFilters = aggregationFilters.concat([`${e.detail.filterOperator}(${refinementString})`]);
          }
        } else {
          aggregationFilters = aggregationFilters.concat(buildFqlRefinementString(selectedFilters));
        }
      } else {
        delete this.searchQuery.requests[0].aggregationFilters;
      }

      if (aggregationFilters.length > 0) {
        this.searchQuery.requests[0].aggregationFilters = aggregationFilters;
      }

      this.resetPagination();

      await this._search(this.searchQuery);
    }
  }

  private async handleSearchInput(e: CustomEvent<ISearchInputEventData>): Promise<void> {
    if (this.shouldRender) {
      // Remove any query string parameter if used as default
      if (this.defaultQueryStringParameter) {
        const url = UrlHelper.removeQueryStringParam(this.defaultQueryStringParameter, window.location.href);
        if (url !== window.location.href) {
          window.history.pushState({}, '', url);
        }
      }

      // If empty keywords, reset to the default state
      const searchKeywords = !isEmpty(e.detail.keywords) ? e.detail.keywords : this.getDefaultQueryText();

      if (searchKeywords && searchKeywords !== this.searchQuery.requests[0].query.queryString) {
        // Update token
        this.tokenService.setTokenValue(BuiltinTokenNames.searchTerms, searchKeywords);

        this.searchQuery.requests[0].query.queryString = searchKeywords;

        this.resetFilters();
        this.resetPagination();

        await this._search(this.searchQuery);
      }
    }
  }

  private async handleSearchVertical(e: CustomEvent<ISearchVerticalEventData>): Promise<void> {
    this.shouldRender = this.selectedVerticalKeys.indexOf(e.detail.selectedVertical.key) !== -1;

    if (this.shouldRender) {
      // Reinitialize search context
      this.resetQueryText();
      this.resetFilters();
      this.resetPagination();
      this.initSortProperties();

      // Update the query when the new tab is selected
      await this._search(this.searchQuery);
    } else {
      // If a query is currently performed, we cancel it to avoid new filters getting populated once one an other tab
      if (this.isLoading) {
        this.msSearchService.abortRequest();
      }

      // Reset available filters for connected search filter components
      this.fireCustomEvent(EventConstants.SEARCH_RESULTS_EVENT, {
        availableFilters: []
      } as ISearchResultsEventData);
    }
  }

  private async handleSearchSort(e: CustomEvent<ISearchSortEventData>): Promise<void> {
    if (this.shouldRender) {
      this.sortProperties = e.detail.sortProperties;

      this.buildSearchQuery();

      // Update the query when new sort is defined
      await this._search(this.searchQuery);
    }
  }

  //#endregion

  //#region Utility methods

  private getDefaultQueryText(): string {
    // 1) Look connected search box if any
    const inputComponent = document.getElementById(this.searchInputComponentId) as any; // MgtSearchInputComponent;
    if (inputComponent && inputComponent.searchKeywords) {
      return inputComponent.searchKeywords;
    }

    // 2) Look query string parameters if any
    if (
      this.defaultQueryStringParameter &&
      !isEmpty(UrlHelper.getQueryStringParam(this.defaultQueryStringParameter, window.location.href))
    ) {
      return UrlHelper.getQueryStringParam(this.defaultQueryStringParameter, window.location.href);
    }

    // 3) Look default hard coded value if any
    if (this.defaultQueryText) {
      return this.defaultQueryText;
    }
  }

  private goToPage(pageNumber: number) {
    if (pageNumber > 0) {
      // "-1" is to calculate the correct index. Ex page "1" with page size "10" means items from start index 0 to index 10.
      this.searchQuery.requests[0].from = (pageNumber - 1) * this.pageSize;
      this._search(this.searchQuery);
    }
  }

  private resetPagination() {
    // Reset to first page
    /* this.searchQuery.requests[0].from = 0;
        const paginationComponent = this.renderRoot.querySelector<MgtPaginationComponent>(`[data-tag-name='${ComponentElements.MgtPaginationElement}']`);
        if (paginationComponent) {
            paginationComponent.initPagination();
        }*/
  }

  private resetFilters() {
    // Reset existing filters if any selected
    /* if (this.searchFiltersComponentId) {
            const filterComponent = document.getElementById(this.searchFiltersComponentId) as MgtSearchFiltersComponent;
            if (filterComponent) {
                filterComponent.clearAllSelectedValues(true);
            }
        }

        delete this.searchQuery.requests[0].aggregationFilters;*/
  }

  private initSortProperties() {
    this.sortProperties = this.sortFieldsConfiguration
      .filter(s => s.isDefaultSort)
      .map(s => {
        return {
          isDescending: s.sortDirection === SortFieldDirection.Descending,
          name: s.sortField
        };
      });
  }

  private resetQueryText() {
    const queryText = this.getDefaultQueryText();
    this.searchQuery.requests[0].query.queryString = queryText;
  }

  /**
   * Notify analytics system
   */
  private trackEvents(results: IMicrosoftSearchDataSourceData) {
    // Search results (except bookmarks)
    if (this.entityTypes.indexOf(EntityType.Bookmark) === -1) {
      const customEventData: ITrackingEventData = {
        action: results.items.length === 0 ? UserAction.SearchResultsNoResult : UserAction.SearchResultsDisplayed,
        category: EventCategory.SearchResultsEvents,
        name: results.items.length.toString()
      };

      if (!this.searchQuery.requests[0].aggregationFilters) {
        // Reset custom value
        customEventData.eventCustomDimensions = [
          {
            key: SearchTrackedDimensions.SelectedFilter,
            value: null
          }
        ];
      }

      this.fireCustomEvent(AnalyticsEventConstants.MONITORED_EVENT, customEventData, true);
    }

    // Bookmarks (in any)
    if (
      this.entityTypes.length === 1 &&
      this.entityTypes.indexOf(EntityType.Bookmark) !== -1 &&
      results.items.length > 0
    ) {
      this.fireCustomEvent(
        AnalyticsEventConstants.MONITORED_EVENT,
        {
          action: UserAction.SearchBookmarksDisplayed,
          category: EventCategory.SearchResultsEvents,
          name: results.items.length.toString()
        } as ITrackingEventData,
        true
      );
    }

    // Search suggestions (if any)
    if (results.queryAlterationResponse?.queryAlteration?.alteredQueryTokens) {
      this.fireCustomEvent(
        AnalyticsEventConstants.MONITORED_EVENT,
        {
          action: UserAction.SearchSuggestionsDisplayed,
          category: EventCategory.SearchResultsEvents,
          name: results.queryAlterationResponse.queryAlteration.alteredQueryTokens.length.toString()
        } as ITrackingEventData,
        true
      );
    }
  }

  //#endregion
}
