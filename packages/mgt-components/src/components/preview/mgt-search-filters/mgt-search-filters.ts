/* eslint-disable max-classes-per-file */
import {
  BuiltinFilterTemplates,
  EventConstants,
  FilterConditionOperator,
  FilterSortDirection,
  FilterSortType,
  IDataFilter,
  IDataFilterConfiguration,
  IDataFilterResult,
  IDataFilterResultValue,
  IDataFilterValue,
  ISearchFiltersEventData,
  ISearchResultsEventData,
  ISearchSortEventData,
  ISearchSortProperty,
  MgtConnectableComponent,
  customElement
} from '@microsoft/mgt-element';
import { isEmpty, cloneDeep, sortBy } from 'lodash-es';
import { strings } from './strings';
import { styles as tailwindStyles } from '../../../styles/tailwind-styles-css';
import { MgtBaseFilterComponent } from './mgt-base-filter';
import { ScopedElementsMixin } from '@open-wc/scoped-elements';
import { MgtCheckboxFilterComponent } from './mgt-checkbox-filter/mgt-checkbox-filter';
import { MgtDateFilterComponent } from './mgt-date-filter/mgt-date-filter';
import { property, state } from 'lit/decorators.js';
import { html, nothing } from 'lit';
import { repeat } from 'lit/directives/repeat.js';

export class MgtSearchFiltersComponentBase extends MgtConnectableComponent {}

@customElement('search-filters')
export class MgtSearchFiltersComponent extends ScopedElementsMixin(MgtSearchFiltersComponentBase) {
  /**
   * The connected search results component ids
   */
  @property({
    type: Array,
    attribute: 'search-results-ids',
    converter: {
      fromAttribute: value => {
        return value.split(',');
      }
    }
  })
  searchResultsComponentIds: string[] = [];

  /**
   * The filters configration
   */
  @property({
    type: Object,
    attribute: 'settings',
    converter: {
      fromAttribute: value => {
        return JSON.parse(value) as IDataFilterConfiguration[];
      }
    }
  })
  filterConfiguration: IDataFilterConfiguration[] = [];

  /**
   * The default logical operatorto use sbetween filters
   */
  @property({ type: String, attribute: 'operator' })
  operator: FilterConditionOperator = FilterConditionOperator.AND;

  /**
   * Available filters received from connected search results component
   */
  @state()
  availableFilters: IDataFilterResult[] = [];

  /**
   * All selected values from all filters combined (not necessarily submitted)
   */
  @state()
  allSelectedFilters: IDataFilter[] = [];

  /**
   * The list of disabled filters
   */
  private declare disabledFilters: string[];

  /**
   * All submitted values from all filters combined
   */
  public declare allSubmittedFilters: IDataFilter[];

  /**
   * The previous applied filters
   */
  private declare previousAvailableFilters: IDataFilterResult[];

  private declare submittedQueryText: string;

  private declare searchResultsEventData: ISearchResultsEventData;
  constructor() {
    super();

    this.allSelectedFilters = [];
    this.allSubmittedFilters = [];
    this.disabledFilters = [];

    this.handleSearchResultsFilters = this.handleSearchResultsFilters.bind(this);
    this.onFilterUpdated = this.onFilterUpdated.bind(this);
    this.onApplyFilters = this.onApplyFilters.bind(this);
    this.onSort = this.onSort.bind(this);
  }

  public render() {
    let renderSort;

    const renderCheckbox = (availableFilter: IDataFilterResult) => {
      return html`<mgt-filter-checkbox
                            class="cursor-pointer"
                            data-ref-name=${availableFilter.filterName}
                            .disabled=${this.disabledFilters.indexOf(availableFilter.filterName) > -1}
                            .filter=${availableFilter}
                            .filterConfiguration=${
                              this.filterConfiguration.filter(c => c.filterName === availableFilter.filterName)[0]
                            }
                            .onFilterUpdated=${this.onFilterUpdated}
                            .onApplyFilters=${this.onApplyFilters}
                        >
                        </mgt-filter-checkbox>`;
    };

    const renderDate = (availableFilter: IDataFilterResult) => {
      return html`<mgt-filter-date
                            class="cursor-pointer"
                            data-ref-name=${availableFilter.filterName}
                            .disabled=${this.disabledFilters.indexOf(availableFilter.filterName) > -1}
                            .filter=${availableFilter}
                            .filterConfiguration=${
                              this.filterConfiguration.filter(c => c.filterName === availableFilter.filterName)[0]
                            }
                            .onFilterUpdated=${this.onFilterUpdated}
                            .onApplyFilters=${this.onApplyFilters}
                        >
                        </mgt-filter-date>`;
    };

    const renderShimmers = html`
            <div class="flex space-x-2">
                ${repeat(
                  this.filterConfiguration,
                  configuration => configuration.filterName,
                  () => {
                    return html`<div class="h-3 w-14 animate-shimmer bg-slate-200 rounded"></div>`;
                  }
                )}
            </div>
        `;

    if (this.searchResultsEventData?.sortFieldsConfiguration && this.availableFilters.length > 0) {
      renderSort = html`
                <ubisoft-search-sort
                    .sortProperties=${this.searchResultsEventData.sortFieldsConfiguration}
                    .onSort=${this.onSort}
                >
                </ubisoft-search-sort>
            `;
    } else {
      renderSort = nothing;
    }

    let renderFilters = html`${repeat(
      // https://lit.dev/docs/templates/lists/#when-to-use-map-or-repeat
      this.availableFilters,
      filter => filter.filterName,
      availableFilter => {
        // Only display the filter component is values are present
        if (availableFilter.values.length > 0) {
          const filterConfiguration = this.getFilterConfiguration(availableFilter.filterName);

          switch (filterConfiguration.template) {
            case BuiltinFilterTemplates.CheckBox:
              return renderCheckbox(availableFilter);

            case BuiltinFilterTemplates.Date:
              return renderDate(availableFilter);

            default:
              return renderCheckbox(availableFilter);
          }
        } else {
          return nothing;
        }
      }
    )}`;

    if (this.availableFilters.length === 0 && this.filterConfiguration.length > 0) {
      if (isEmpty(this.submittedQueryText)) {
        renderFilters = renderShimmers;
      } else {
        renderFilters = html`<span>${strings.noFilters}</span>`;
      }
    }

    return html`
            <div class="px-2.5">
                <div class="max-w-7xl ml-auto mr-auto mb-8">
                    <div class="font-sans text-sm flex py-[16px] px-[32px] rounded-lg shadow-filtersShadow bg-light300 justify-between">
                        <div class="flex flex-wrap items-center space-x-2">
                            <egg-icon icon-id="egg-global:action:filter-range"></egg-icon>
                            ${renderFilters}
                            ${
                              (this.availableFilters.length > 0 && this.allSelectedFilters.length > 0) ||
                              this.allSubmittedFilters.length > 0
                                ? html`<button data-ref="reset" class="flex cursor-pointer space-x-1 items-center hover:text-primary opacity-75" @click=${() => {
                                    this.clearAllSelectedValues();
                                  }}>
                                        <egg-icon icon-id="egg-global:action:restore"></egg-icon>
                                        <span>${strings.resetAllFilters}</span>
                                    </button>`
                                : null
                            }
                        </div>
                        <div>
                            ${renderSort}
                        </div>
                    </div>
                </div>
            </div>
        `;
  }

  static get styles() {
    return [tailwindStyles];
  }

  static get scopedElements() {
    return {
      'mgt-filter-checkbox': MgtCheckboxFilterComponent,
      'mgt-filter-date': MgtDateFilterComponent
      //"mgt-search-sort": MgtSearchSortComponent
    };
  }

  protected get strings(): { [x: string]: string } {
    return strings;
  }

  public connectedCallback(): Promise<void> {
    const bindings = this.searchResultsComponentIds.map(componentId => {
      return {
        id: componentId,
        eventName: EventConstants.SEARCH_RESULTS_EVENT,
        callbackFunction: this.handleSearchResultsFilters
      };
    });

    this.bindComponents(bindings);

    return Promise.resolve(super.connectedCallback());
  }

  public handleSearchResultsFilters(e: CustomEvent<ISearchResultsEventData>) {
    this.searchResultsEventData = e.detail;
    this.availableFilters = cloneDeep(e.detail.availableFilters);
    this.submittedQueryText = e.detail.submittedQueryText;

    // Handle filters with zero value
    if (e.detail.availableFilters.length === 0 && this.allSelectedFilters.length > 0) {
      // Existing selected filters
      const selectedFilterNames = this.allSelectedFilters.map(filter => filter.filterName);

      this.availableFilters = cloneDeep(
        this.previousAvailableFilters.map(f => {
          f.values = [];
          return f;
        })
      );

      // Disable non selected filters
      this.previousAvailableFilters.forEach(p => {
        if (selectedFilterNames.indexOf(p.filterName) === -1) {
          this.disabledFilters.push(p.filterName);
        }
      });
    } else {
      this.disabledFilters = [];
    }

    // Merge filters with the same field name
    // This scenario happens when multiple sources have the same alias as refiner. In this case, the API returns duplicate fields instead of merging them.
    this.availableFilters = this.mergeFilters(this.availableFilters);

    // Sort filters by configuration if any
    this.availableFilters = this.availableFilters.sort((a, b) => {
      const aSortIdx = this.filterConfiguration.filter(f => f.filterName === a.filterName)[0].sortIdx;
      const bSortIdx = this.filterConfiguration.filter(f => f.filterName === b.filterName)[0].sortIdx;
      return aSortIdx - bSortIdx;
    });

    // Save available filters for subsequent usage
    this.previousAvailableFilters = cloneDeep(this.availableFilters);
  }

  /**
   * Handler when a value is updated (selected/unselected)from a specific filter
   * @param filterName the filter name from where values has been applied
   * @param filterValue the filter value that has been updated
   */
  private onFilterUpdated(filterName: string, filterValue: IDataFilterValue, selected: boolean) {
    if (selected) {
      // Get the index of the filter in the current selected filters collection
      const filterIdx = this.allSelectedFilters
        .map(filter => {
          return filter.filterName;
        })
        .indexOf(filterName);
      const newFilters = [...this.allSelectedFilters];

      if (filterIdx !== -1) {
        const valueIdx = this.allSelectedFilters[filterIdx].values.map(v => v.key).indexOf(filterValue.key);

        if (valueIdx === -1) {
          // If the value does not exist yet, we add it to the selected values

          newFilters[filterIdx].values.push(filterValue);
          this.allSelectedFilters = newFilters;
        } else {
          // Otherwise, we update the value in selected values
          newFilters[filterIdx].values[valueIdx] = filterValue;
        }
      } else {
        const newFilter: IDataFilter = {
          filterName: filterName,
          values: [filterValue]
        };

        newFilters.push(newFilter);
      }

      this.allSelectedFilters = newFilters;
    } else {
      // Remove the filter value
      this.allSelectedFilters = this.allSelectedFilters.map(selectedFilter => {
        selectedFilter.values = selectedFilter.values.filter(value => value.key != filterValue.key);

        if (selectedFilter.values.length > 0) {
          return selectedFilter;
        } else {
          return null;
        }
      });

      // Remove null values
      this.allSelectedFilters = this.allSelectedFilters.filter(f => f);
    }
  }

  /**
   * Handler when values are applied from a specific filter
   * @param filterName the filter name from where values has been applied
   */
  private onApplyFilters(filterName: string) {
    // Update the list of submitted filters to sent to the search engine
    const selectedFilter = this.allSelectedFilters.filter(f => f.filterName === filterName)[0];

    // If filter is 'null', it means no values are currently selected for that filter
    if (selectedFilter) {
      const filterIdx = this.allSubmittedFilters
        .map(filter => {
          return filter.filterName;
        })
        .indexOf(filterName);
      const newFilters = [...this.allSubmittedFilters];

      if (filterIdx === -1) {
        newFilters.push(selectedFilter);
      } else {
        newFilters[filterIdx] = selectedFilter;
      }

      this.allSubmittedFilters = newFilters;
    } else {
      this.allSubmittedFilters = this.allSubmittedFilters.filter(s => s.filterName !== filterName);
    }

    // Reset selected values from other filters that are not already submitted
    const otherFiltersWithSelectedValues = this.allSelectedFilters
      .filter(f => {
        return (
          f.filterName !== filterName &&
          this.allSubmittedFilters.filter(a => a.filterName === f.filterName).length === 0
        );
      })
      .map(f => f.filterName);

    otherFiltersWithSelectedValues.forEach(filterName => {
      const filterComponent = this.getFilterComponents(filterName);
      if (filterComponent) {
        filterComponent[0].clearSelectedValues(true);
      }
    });

    this.applyFilters();
  }

  /**
   * Handler when sort properties are updated
   * @param sortProperties the sort properties
   */
  private onSort(sortProperties: ISearchSortProperty[]) {
    this.fireCustomEvent(
      EventConstants.SEARCH_SORT_EVENT,
      {
        sortProperties: sortProperties
      } as ISearchSortEventData,
      true
    );
  }

  /**
   * Send filters to connected search results component
   */
  private applyFilters() {
    // Update filters information before sending them to source so they can be processed according their specificities
    // i.e Merge selected filter with relavant information from its configuration
    const updatedFilters = this.allSubmittedFilters.map(f => {
      const confguration = this.getFilterConfiguration(f.filterName);
      if (confguration) {
        f.operator = confguration.operator;
      }
      return f;
    });

    this.fireCustomEvent(EventConstants.SEARCH_FILTER_EVENT, {
      selectedFilters: updatedFilters,
      filterOperator: this.operator ? this.operator : FilterConditionOperator.AND
    } as ISearchFiltersEventData);
  }

  private getFilterConfiguration(filterName: string): IDataFilterConfiguration {
    return this.filterConfiguration.filter(c => c.filterName === filterName)[0];
  }

  /**
   * Merges filter values having the same filter name
   * @param availableFilters the available filters returned from the search response
   * @returns the merged filters
   */
  private mergeFilters(availableFilters: IDataFilterResult[]): IDataFilterResult[] {
    let allMergedFilters: IDataFilterResult[] = [];

    availableFilters.forEach(filterResult => {
      const mergedFilterIdx = allMergedFilters.map(m => m.filterName).indexOf(filterResult.filterName);

      if (mergedFilterIdx === -1) {
        allMergedFilters.push(filterResult);
      } else {
        const allMergedValues: IDataFilterResultValue[] = [];
        const allValues = allMergedFilters[mergedFilterIdx].values.concat(filterResult.values);

        // 3. Sum counts for similar value names
        allValues.forEach(value => {
          const mergedValueIdx = allMergedValues.map(v => v.name).indexOf(value.name);

          if (mergedValueIdx === -1) {
            allMergedValues.push(value);
          } else {
            allMergedValues[mergedValueIdx].count = allMergedValues[mergedValueIdx].count + value.count;
          }
        });

        allMergedFilters[mergedFilterIdx].values = allMergedValues;
      }
    });

    // Sort values according to the filter configuration
    allMergedFilters = allMergedFilters.map(filter => {
      let sortByField = 'name';
      let sortDirection = FilterSortDirection.Ascending;

      const filterConfigurationIdx = this.filterConfiguration
        .map(configuration => configuration.filterName)
        .indexOf(filter.filterName);
      if (filterConfigurationIdx !== -1) {
        const filterConfiguration = this.filterConfiguration[filterConfigurationIdx];
        if (filterConfiguration.sortBy === FilterSortType.ByCount) {
          sortByField = 'count';
        }

        if (filterConfiguration.sortDirection === FilterSortDirection.Descending) {
          sortDirection = FilterSortDirection.Descending;
        }
      }

      filter.values =
        sortDirection === FilterSortDirection.Ascending
          ? sortBy(filter.values, sortByField)
          : sortBy(filter.values, sortByField).reverse();

      return filter;
    });

    return allMergedFilters;
  }

  public clearAllSelectedValues(preventApply?: boolean) {
    // Reset selected values for all child components
    const filterComponents = this.getFilterComponents();

    // Reset all filters from the UI (whithout submitting values)
    filterComponents.forEach(component => component.clearSelectedValues(true));

    if (this.allSubmittedFilters.length > 0) {
      this.allSubmittedFilters = [];

      // Apply empty filters
      if (!preventApply) {
        this.applyFilters();
      }
    }
  }

  /**
   * Retrieved the list of child filters
   * @param filterName Optionnal. A specific filter name to retrieve
   * @returns the list of child filter components
   */
  private getFilterComponents(filterName?: string): MgtBaseFilterComponent[] {
    // Reset all values from sub components. List all valid sub filter components
    const filterComponents: MgtBaseFilterComponent[] = Array.prototype.slice.call(
      this.renderRoot.querySelectorAll<MgtBaseFilterComponent>(`
            [data-tag-name='mgt-filter-date'],
            [data-tag-name='mgt-filter-checkbox']
        `)
    );

    return filterName
      ? filterComponents.filter(component => component.filter.filterName === filterName)
      : filterComponents;
  }
}
