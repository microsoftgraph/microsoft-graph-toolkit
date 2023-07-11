/* eslint-disable @typescript-eslint/no-unsafe-call */
/* eslint-disable @typescript-eslint/no-unsafe-member-access */
/* eslint-disable @typescript-eslint/no-unsafe-assignment */
/* eslint-disable @typescript-eslint/no-unsafe-argument */
/* eslint-disable @typescript-eslint/unbound-method */
/* eslint-disable @typescript-eslint/no-unnecessary-type-assertion */
import { html, PropertyValues } from 'lit';
import { cloneDeep } from 'lodash-es';
import { repeat } from 'lit/directives/repeat.js';
import { unsafeHTML } from 'lit/directives/unsafe-html.js';
import { strings } from './strings';
import { MgtBaseFilterComponent } from '../mgt-base-filter';
import { IDataFilterResultValue, IDataFilterAggregation } from '@microsoft/mgt-element';
import { state } from 'lit/decorators.js';
import { SvgIcon, getSvg } from '../../../../utils/SvgHelper';
import { getFileTypeIconUriByExtension } from '../../../../styles/fluent-icons';

export class MgtCheckboxFilterComponent extends MgtBaseFilterComponent {
  @state()
  searchKeyword: string;

  /**
   * List of filtered values
   */
  @state()
  filteredValues: IDataFilterResultValue[] = [];

  declare startOffset: number;

  /**
   * Number of items to be displayed in the menu. Limit this number to increase performances
   */
  declare pageSize: number;

  constructor() {
    super();

    this.startOffset = 0;
    this.pageSize = 50;

    this.onScroll = this.onScroll.bind(this);
  }

  public renderFilterContent() {
    // Display only filter values submitted and present in the available filter values
    // Multiple refinement steps can lead initial selected values to not be included in the available values
    const selectedValues = this.submittedFilterValues.filter(s => {
      return this.filter.values.map(v => v.key).indexOf(s.key) !== -1;
    });

    const filterName = this.localizedFilterName ? this.localizedFilterName.toLowerCase() : null;

    const renderSearchBox = html`
            <div class="relative">
                <div class="absolute top-[10px] left-[10px] text-black opacity-25">${getSvg(SvgIcon.Search)}</div>
                
                ${
                  this.searchKeyword
                    ? html`
                      <div class="absolute top-[10px] right-[10px] text-black opacity-25" @click=${
                        this.clearSearchKeywords
                      }>${getSvg(SvgIcon.Close)}</div>`
                    : null
                }

                <fluent-text-field  id="searchbox"
                        autocomplete="off"
                        class="w-full outline-none bg-transparent" 
                        type="text" 
                        placeholder=${strings.searchPlaceholder}
                        @input=${e => {
                          this.filterValues(e.target.value);
                        }}
                ></fluent-text-field>
                </fluent-text-field>
            </div>
        `;

    return html`
                <div class="sticky top-0 flex flex-col space-y-2 z-10"
    
                          @mousedown=${(e: Event) => {
                            e.preventDefault();
                            // e.stopPropagation();
                          }} 
                > 
                    <div class="border-b px-6 py-3 space-y-2">                  
                        <label class="text-base">${this.filter.values.length} ${filterName}</label>
                        ${renderSearchBox}                         
                    </div>

                    <div class="flex justify-between items-center px-6 py-3 min-h-[48px]">

                        <div class="opacity-75"><label>${selectedValues.length} ${strings.selections}</label></div>
                        ${
                          this.selectedValues.length > 0 || this.submittedFilterValues.length > 0
                            ? html`
                              <button data-ref="reset" type="reset" class="flex cursor-pointer space-x-1 items-center hover:text-primary opacity-75" @click=${() =>
                                this.clearSelectedValues()}>
                                    <div>${getSvg(SvgIcon.Refresh)}</div>
                                    <span>${strings.reset}</span>
                                </button>`
                            : null
                        }
                    </div>
                </div>
                <div 
                    id="filter-menu-content"
                    class="p-1 flex-col flex space-y-1">
                    ${repeat(
                      this.filteredValues,
                      filterValue => filterValue.key,
                      filterValue => {
                        return html`
                            <fluent-option
                                class="hover:rounded-lg p-2 block"
                                @click=${() => {
                                  this.onItemUpdated(filterValue, !this.isSelectedValue(filterValue.key));

                                  if (!this.filterConfiguration.isMulti) {
                                    this.applyFilters();
                                  }
                                }}
                                data-ref-value=${filterValue.value}
                                value=${filterValue.value}
                                data-ref-name=${filterValue.name}
                                .selected=${this.isSelectedValue(filterValue.key)}
                            >
                                <div class="flex items-center space-x-2">
                                    <div class="flex items-center space-x-2">
                                        ${
                                          this.getFilterAggregation(filterValue.name)
                                            ? html`
                                                <img  data-ref="icon" 
                                                      class="w-[24px]" 
                                                      src=${getFileTypeIconUriByExtension(
                                                        this.getFilterAggregation(filterValue.name)
                                                          .aggregationValueIconUrl,
                                                        48,
                                                        'svg'
                                                      )}
                                                />`
                                            : null
                                        }
                                        <div data-ref="value" class="font-medium ${
                                          this.isSelectedValue(filterValue.key) ? 'text-primary' : ''
                                        }">
                                            ${
                                              this.searchKeyword
                                                ? this.highlightMatches(filterValue.name)
                                                : filterValue.name
                                            }
                                        </div>
                                    </div>
                                    ${
                                      this.filterConfiguration.showCount
                                        ? html`
                                            <div>
                                                <label class="opacity-75">(${filterValue.count})</label>
                                            </div>
                                        `
                                        : null
                                    }
                                </div>
                              </fluent-option>
                            `;
                      }
                    )}
                </div>
                ${
                  this.filterConfiguration.isMulti
                    ? html`
                        <div 
                          class="sticky bottom-0 flex justify-around py-2 px-2 space-x-4 w-full border-opacity-25"
                          @mousedown=${(e: Event) => {
                            e.preventDefault();
                            e.stopPropagation();
                          }} 
                        >
                            <fluent-button data-ref="cancel" type="submit" class="flex items-center justify-center text-primary font-medium w-24" @click=${
                              this.closeMenu
                            }>
                                ${strings.cancel}
                            </fluent-button>
                            <fluent-button data-ref="apply" type="submit" appearance="accent" class="flex items-center justify-center text-primary font-medium w-24 ${
                              this.selectedValues.length === 0 || !this.canApplyValues
                                ? 'opacity-50 cursor-not-allowed'
                                : ''
                            }"
                                    ?disabled=${this.selectedValues.length === 0 || !this.canApplyValues}
                                    @click=${this.applyFilters}
                            >
                                ${strings.apply}
                            </fluent-button>
                        </div>
                    `
                    : null
                }
        `;
  }

  public firstUpdated(changedProperties: PropertyValues<this>): void {
    // Return only a subset of values to manage performances
    this.filter.values = this.processAggregations(this.filter.values);
    this.filteredValues = cloneDeep(this.filter.values).slice(0, this.pageSize);

    this.startOffset = this.pageSize;

    const elt = this.renderRoot.querySelector('#filter-menu-content');
    elt.addEventListener('scroll', this.onScroll);

    super.firstUpdated(changedProperties);
  }

  public updated(changedProperties: PropertyValues<this>): void {
    if (changedProperties.get('filter')) {
      this.filter.values = this.processAggregations(this.filter.values);
      this.filteredValues = this.sortBySelectedValues(cloneDeep(this.filter.values)).slice(0, this.pageSize);
    }

    super.updated(changedProperties);
  }

  public disconnectedCallback(): void {
    const elt = this.renderRoot.querySelector('#filter-menu-content');
    elt.removeEventListener('scroll', this.onScroll);
    super.disconnectedCallback();
  }

  protected get strings(): { [x: string]: string } {
    return strings;
  }

  public filterValues(value: string) {
    if (!value) {
      this.filteredValues = this.sortBySelectedValues(cloneDeep(this.filter.values));
    } else {
      this.filteredValues = this.filter.values.filter(v => v.name.toLocaleLowerCase().indexOf(value) !== -1);
    }

    // Return only a subset of values to manage performances
    this.filteredValues = this.filteredValues.slice(0, this.pageSize);

    this.searchKeyword = value;
  }

  private sortBySelectedValues(filters: IDataFilterResultValue[]) {
    return filters.sort((x, y) => {
      return Number(this.isSelectedValue(y.key)) - Number(this.isSelectedValue(x.key));
    });
  }

  private highlightMatches(value: string) {
    const matchExpr = value.replace(
      new RegExp(this.searchKeyword, 'gi'),
      match => `<b class="text-primary opacity-50">${match}</b>`
    );
    return html`${unsafeHTML(matchExpr)}`;
  }

  private clearSearchKeywords() {
    (this.renderRoot.querySelector('#searchbox') as HTMLInputElement).value = null;
    this.filterValues(null);
  }

  private onScroll() {
    const elt = this.renderRoot.querySelector('#filter-menu-content');
    if (elt.scrollTop + elt.clientHeight >= elt.scrollHeight - 50) {
      const newOffset = this.startOffset + this.pageSize;
      this.filteredValues = this.filter.values.slice(0, newOffset);
      this.startOffset = newOffset;
    }
  }

  private getFilterAggregation(name: string): IDataFilterAggregation {
    return this.filterConfiguration.aggregations?.filter(aggregation => {
      return this.getLocalizedString(aggregation.aggregationName) === name;
    })[0];
  }
}
