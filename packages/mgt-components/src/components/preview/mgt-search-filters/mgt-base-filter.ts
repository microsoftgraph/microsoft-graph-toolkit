import { MgtConnectableComponent } from '@microsoft/mgt-element';
import { html, PropertyValues, TemplateResult } from 'lit';
import { state, property } from 'lit/decorators.js';
import { isEqual, cloneDeep, sumBy, orderBy } from 'lodash-es';
import { IDataFilterResult, IDataFilterValue, IDataFilterResultValue } from '@microsoft/mgt-element';
import { IDataFilterConfiguration, FilterSortType, FilterSortDirection } from '@microsoft/mgt-element';
import { fluentSelect, fluentOption, provideFluentDesignSystem, fluentListbox } from '@fluentui/web-components';
import { styles as tailwindStyles } from '../../../styles/tailwind-styles-css';

export enum DateFilterKeys {
  From = 'from',
  To = 'to'
}

export abstract class MgtBaseFilterComponent extends MgtConnectableComponent {
  /**
   * Filter information to display
   */
  @property()
  filter: IDataFilterResult;

  /**
   * Filter confguration
   */
  @property()
  filterConfiguration: IDataFilterConfiguration;

  /**
   * Flag indicating if the filter should be disabled
   */
  @property()
  disabled: boolean;

  /**
   * Callback function when a filter is selected
   */
  @property()
  onFilterUpdated: (filterName: string, filterValue: IDataFilterValue, selected: boolean) => void;

  /**
   * Callback function when filters are submitted
   */
  @property()
  onApplyFilters: (filterName: string) => void;

  @state()
  isExpanded: boolean;

  /**
   * The current selected values in the component
   */
  @state()
  selectedValues: IDataFilterValue[] = [];

  /**
   * The submitted filter values
   */
  public declare submittedFilterValues: IDataFilterValue[];

  /**
   * Mutation observer for the fitler button
   */
  declare buttonObserver: MutationObserver;

  protected get localizedFilterName(): string {
    return this.getLocalizedString(this.filterConfiguration.displayName);
  }

  /**
   * Flag indicating if the selected values can be applied as filters
   */
  protected get canApplyValues(): boolean {
    return !isEqual(this.submittedFilterValues.map(v => v.value).sort(), this.selectedValues.map(v => v.value).sort());
  }

  public constructor() {
    super();

    this.submittedFilterValues = [];

    this.onItemUpdated = this.onItemUpdated.bind(this);
    this.applyFilters = this.applyFilters.bind(this);
    this.closeMenu = this.closeMenu.bind(this);

    // Register fluent tabs (as scoped elements)
    provideFluentDesignSystem().register(fluentSelect(), fluentOption(), fluentListbox());
  }

  public render() {
    let renderFilterName = html`<span data-ref-name=${this.localizedFilterName}>${this.localizedFilterName}</span>`;

    if (this.submittedFilterValues.length === 1) {
      renderFilterName = html`<span class="font-bold">${this.submittedFilterValues[0].name}</span>`;
    }

    if (this.submittedFilterValues.length > 1) {
      // Display only filter values submitted and present in the available filter values
      // Multiple refinement steps can lead initial selected values to not be included in the available values
      const selectedValues = this.submittedFilterValues.filter(s => {
        return (
          this.filter.values.map(v => v.key).indexOf(s.key) !== -1 ||
          s.key === DateFilterKeys.From ||
          s.key === DateFilterKeys.To
        );
      });

      renderFilterName = html`
                <div class="flex items-center space-x-2">
                    <div class="font-bold">${this.localizedFilterName}</div>
                    <div class="rounded-[50%] flex items-center justify-center w-[22px] h-[22px] bg-primary font-bold text-white ">${selectedValues.length}</div>
                </div>
            `;
    }

    return html`

                <fluent-select title=${this.localizedFilterName}>
                  <div slot="selected-value">
                    ${renderFilterName}
                  </div>
                 
                <div @click=${(e: Event) => {
                  e.stopPropagation();
                }}>
                            ${this.renderFilterContent()}
                        </div>            
                </fluent-select>
                `;
  }

  protected firstUpdated(changedProperties: PropertyValues<this>): void {
    // Set the element id to uniquely identify it in the DOM
    this.id = this.filter.filterName;

    const filterButton = this.renderRoot.querySelector("[data-tag-name='egg-button']");

    this.buttonObserver = new MutationObserver(_mutations => {
      _mutations.forEach(mutation => {
        switch (mutation.type) {
          case 'attributes':
            if (mutation.attributeName === 'aria-expanded') {
              this.isExpanded = (mutation.target as HTMLElement).getAttribute('aria-expanded') === 'true';
            }
            break;
          default:
            break;
        }
      });
    });

    if (filterButton) {
      this.buttonObserver.observe(filterButton, {
        attributeFilter: ['aria-expanded'],
        attributeOldValue: true,
        childList: false,
        subtree: false
      });
    }

    super.firstUpdated(changedProperties);
  }

  public disconnectedCallback(): void {
    // Could be already disconnected by an other filter instance
    if (this.buttonObserver) {
      this.buttonObserver.disconnect();
    }

    super.disconnectedCallback();
  }

  /**
   * Reset all selected values for the current filter in the UI.
   * Can be called from parent components
   */
  protected resetSelectedValues() {
    // Update parent state
    this.selectedValues.forEach(v => {
      this.onFilterUpdated(this.filter.filterName, v, false);
    });

    // Reset internal state
    let newValues = [...this.selectedValues];
    newValues = [];

    this.selectedValues = newValues;
  }

  /**
   * Clear all selected values in the UI and submit empty filters to connected components
   * @param preventApply Set true if you want to prevent new values to be submitted for that filter
   */
  public clearSelectedValues(preventApply?: boolean) {
    this.resetSelectedValues();

    if (this.submittedFilterValues.length > 0) {
      // Reset submitted filters
      this.submittedFilterValues = [];

      if (!preventApply) this.applyFilters();
    }
  }

  protected onItemUpdated(filterValue: IDataFilterValue, selected: boolean) {
    if (selected) {
      // Get the index of the filter value in the current selected values collection
      const valueIdx = this.selectedValues.map(value => value.key).indexOf(filterValue.key);
      const newValues = [...this.selectedValues];

      if (valueIdx !== -1) {
        // Update the existing value
        newValues[valueIdx] = filterValue;
      } else {
        // Add the new value if doesn't exist
        newValues.push(filterValue);
      }

      this.selectedValues = newValues;
    } else {
      this.selectedValues = this.selectedValues.filter(v => v.key !== filterValue.key);
    }

    this.onFilterUpdated(this.filter.filterName, filterValue, selected);
  }

  protected isSelectedValue(key: string): boolean {
    const isSelected =
      this.selectedValues.filter(v => {
        return v.key === key;
      }).length > 0;

    return isSelected;
  }

  /**
   * The filter content to be implemented by concrete classes
   */
  protected abstract renderFilterContent(): TemplateResult;

  protected applyFilters(): void {
    this.submittedFilterValues = cloneDeep(this.selectedValues);
    this.onApplyFilters(this.filter.filterName);
    this.closeMenu();
  }

  /**
   * Process manual filter aggregations according to matched values from configuration
   * @param values the original filter values received from the results
   * @returns the new aggregated filters values
   */
  protected processAggregations(values: IDataFilterResultValue[]): IDataFilterResultValue[] {
    let filteredValues = cloneDeep(values);

    if (this.filterConfiguration.aggregations) {
      this.filterConfiguration.aggregations.forEach(aggregation => {
        // Get all matching values
        const matchingValues = filteredValues.filter(value => {
          return aggregation.matchingValues.indexOf(value.name) > -1;
        });

        // Remove all values matching the aggregation
        filteredValues = filteredValues.filter(value => {
          return aggregation.matchingValues.indexOf(value.name) === -1;
        });

        if (matchingValues.length > 0) {
          // A new aggregation
          filteredValues.push({
            count: sumBy(matchingValues, 'count'),
            key: this.getLocalizedString(aggregation.aggregationName),
            name: this.getLocalizedString(aggregation.aggregationName),
            value: aggregation.aggregationValue
          });
        }
      });
    }

    // Sort values according to the filter configuration
    const sortProperty = this.filterConfiguration.sortBy === FilterSortType.ByCount ? 'count' : 'name';
    const sortDirection = this.filterConfiguration.sortDirection === FilterSortDirection.Ascending ? 'asc' : 'desc';
    filteredValues = orderBy(filteredValues, sortProperty, sortDirection);

    return filteredValues;
  }

  // Close filter menu
  public closeMenu() {
    if (this.selectedValues.length > 0 && this.submittedFilterValues.length === 0) {
      this.resetSelectedValues();
    }

    const eggMenu = this.renderRoot.querySelector("[data-tag-name='egg-menu']");
    eggMenu.removeAttribute('opened');
  }

  static get styles() {
    return [tailwindStyles];
  }
}
