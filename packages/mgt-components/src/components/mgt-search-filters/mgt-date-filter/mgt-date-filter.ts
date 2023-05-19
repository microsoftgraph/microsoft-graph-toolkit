import { DateHelper, FilterComparisonOperator, IDataFilterValue, LocalizationHelper } from '@microsoft/mgt-element';
import { customElement, html, PropertyValues } from 'lit-element';
import { nothing } from 'lit-html';
import { MgtBaseFilterComponent, DateFilterKeys } from '../mgt-base-filter';
import { strings } from './strings';

export enum DateFilterInterval {
  AnyTime,
  Today,
  Past24,
  PastWeek,
  PastMonth,
  Past3Months,
  PastYear,
  OlderThanAYear
}

export class MgtDateFilterComponent extends MgtBaseFilterComponent {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private declare dayJs: any;
  private declare dateHelper: DateHelper;

  private declare allIntervals: { [key in DateFilterInterval]: string };

  public get fromDate(): string {
    return this.getDateValue(DateFilterKeys.From);
  }

  public get toDate(): string {
    return this.getDateValue(DateFilterKeys.To);
  }

  public constructor() {
    super();

    this.allIntervals = this.getAllIntervals();

    this.dateHelper = new DateHelper(LocalizationHelper.strings?.language);

    this.onUpdateFromDate = this.onUpdateFromDate.bind(this);
    this.onUpdateToDate = this.onUpdateToDate.bind(this);
    this.applyDateFilters = this.applyDateFilters.bind(this);
  }

  public renderFilterContent() {
    // Intervals to display
    const intervals = this.filter.values.filter(v => {
      // If an interval is already selected, limit to only one to avoid selecting multiple matched intervals for the same items
      // Ex: an item created last month will also match the past 3 months, and past year intervals. However, it does not make sense to propose all at once once user selects one.
      if (this.selectedValues.length > 0) {
        return this.selectedValues.some(s => s.key === v.key);
      }
      return v.count > 0;
    });

    return html`
                    <div class="sticky top-0 flex justify-between items-center px-6 py-3 space-x-2 bg-white z-10 min-h-[48px]">
                        <div class="opacity-75"><label>${this.selectedValues.length} ${strings.selections}</label></div>
                        ${
                          this.selectedValues.length > 0
                            ? html`<div class="flex cursor-pointer space-x-1 items-center hover:text-primary opacity-75" @click=${() =>
                                this.clearSelectedValues()}>
                                    <egg-icon icon-id="egg-global:action:restore"></egg-icon>
                                    <span>${strings.reset}</span>
                                </div>`
                            : null
                        }
                    </div>
                    <div class="p-2">
                        ${
                          !this.fromDate && !this.toDate
                            ? intervals.map(filterValue => {
                                return html`
                                        <egg-menu-item
                                            class="hover:rounded-lg" 
                                            @click=${() => {
                                              this.onItemUpdated(filterValue, !this.isSelectedValue(filterValue.key));

                                              if (!this.filterConfiguration.isMulti) {
                                                // Apply filters immediately
                                                this.applyFilters();
                                              }
                                            }} 
                                        >
                                            <div class="flex items-center space-x-3">
                                                <div class="place-items-center">
                                                    <input  type="checkbox" 
                                                            value="${filterValue.value}" 
                                                            .checked=${this.isSelectedValue(filterValue.key)}
                                                            class="checked:bg-primary focus:opacity-75 text-primary rounded-sm border border-opacity-25 focus:ring-primary"
                                                    >
                                                </div>
                                                <div>
                                                    <label class="font-medium ${
                                                      this.isSelectedValue(filterValue.key) ? 'text-primary' : ''
                                                    }">${this._getIntervalForValue(filterValue)}</label>
                                                </div>
                                                ${
                                                  this.filterConfiguration.showCount
                                                    ? html`<div>
                                                            <label class="opacity-75">(${filterValue.count})</label>
                                                        </div>
                                                    `
                                                    : null
                                                }
                                            </div>
                                        </egg-menu-item>
                                `;
                              })
                            : nothing
                        }
                    </div>
                    <div class="flex flex-col px-6 py-3 text-sm space-y-2 bg-white border-t border-gray-400 border-opacity-25">
                        <div class="flex flex-col">
                            <span class="opacity-75">${strings.from}:</span>
                            <input  id="from" 
                                    type="date"
                                    max=${this.toDate} 
                                    .value=${this.fromDate} 
                                    @change=${this.onUpdateFromDate}  
                            />   
                        </div>
                        
                        <div class="flex flex-col">
                            <span class="opacity-75">${strings.to}:</span>
                            <input  id="to"
                                    type="date"
                                    min=${this.fromDate}
                                    max=${new Date().toISOString().split('T')[0]}
                                    .value=${this.toDate} 
                                    @change=${this.onUpdateToDate}
                            />
                        </div>                        
                        <button class="flex justify-center text-white rounded-lg bg-primary px-4 py-1 min-w-[140px] font-medium ${
                          (!this.toDate && !this.fromDate) || !this.canApplyValues
                            ? 'opacity-50 cursor-not-allowed pointer-events-none'
                            : ''
                        }" @click=${this.applyDateFilters}>${strings.applyDates}</button>
                    </div>
            `;
  }

  public async connectedCallback(): Promise<void> {
    this.dayJs = await this.dateHelper.dayJs();
    super.connectedCallback();
  }

  protected get strings(): { [x: string]: string } {
    return strings;
  }

  protected updated(changedProperties: PropertyValues<this>): void {
    this.allIntervals = this.getAllIntervals();
    super.updated(changedProperties);
  }

  private _getIntervalDate(unit: string, count: number): Date {
    return this._getIntervalDateFromStartDate(new Date(), unit, count);
  }

  private _getIntervalDateFromStartDate(startDate: Date, unit: string, count: number): Date {
    return this.dayJs(startDate).subtract(count, unit);
  }

  private _getIntervalForValue(filterValue: IDataFilterValue): string {
    const dateAsString = filterValue.value;

    if (dateAsString && this.dayJs) {
      let dateRanges = [];
      // Value from Microaoft Search for date properties as FQL filter value
      if (dateAsString.indexOf('range(') !== -1) {
        const matches = dateAsString.match(
          /(min|max)|(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2}(?:\.\d*)?)((-(\d{2}):(\d{2})|Z)?)/gi
        );
        if (matches) {
          // Return date range (i.e. dates between parenthesis)
          dateRanges = matches;
        }
      }

      // To get it work, we need to submit equivalent aggregations at query time
      const past24Date = this._getIntervalDate('days', 1);
      const pastWeekDate = this._getIntervalDate('weeks', 1);
      const pastMonthDate = this._getIntervalDate('months', 1);
      const past3MonthsDate = this._getIntervalDate('months', 3);
      const pastYearDate = this._getIntervalDate('years', 1);

      // Mutate the original object to get the correct name when submitted
      if (dateRanges.indexOf('min') !== -1) {
        filterValue.name = this.allIntervals[DateFilterInterval.OlderThanAYear];
      } else if (dateRanges.indexOf('max') !== -1) {
        filterValue.name = this.allIntervals[DateFilterInterval.Today];
      } else if (this.dayJs(dateRanges[0]).isSame(past24Date, 'day')) {
        filterValue.name = this.allIntervals[DateFilterInterval.Past24];
      } else if (this.dayJs(dateRanges[0]).isSame(pastWeekDate, 'day')) {
        filterValue.name = this.allIntervals[DateFilterInterval.PastWeek];
      } else if (this.dayJs(dateRanges[0]).isSame(pastMonthDate, 'day')) {
        filterValue.name = this.allIntervals[DateFilterInterval.PastMonth];
      } else if (this.dayJs(dateRanges[0]).isSame(past3MonthsDate, 'day')) {
        filterValue.name = this.allIntervals[DateFilterInterval.Past3Months];
      } else if (this.dayJs(dateRanges[0]).isSame(pastYearDate, 'day')) {
        filterValue.name = this.allIntervals[DateFilterInterval.PastYear];
      }
    } else {
      filterValue.name = this.allIntervals[DateFilterInterval.AnyTime];
    }

    return filterValue.name;
  }

  private onUpdateFromDate(e: InputEvent) {
    let date = (e.target as HTMLInputElement).value;

    // In the case user enter manually a date later than today
    if (this.dayJs(date).isValid() && this.dayJs(date).isAfter(this.dayJs())) {
      date = new Date().toISOString().split('T')[0];
    }

    const filterName = `${strings.from} ${this.dayJs(date).format('ll')}`;
    this.onItemUpdated(
      {
        key: DateFilterKeys.From,
        name: filterName,
        value: date ? new Date(date).toISOString() : null,
        operator: FilterComparisonOperator.Geq
      },
      !!date // No value = unselected
    );

    if (!date && this.submittedFilterValues.length > 0) {
      this.applyFilters();
    }
  }

  private onUpdateToDate(e) {
    let date = (e.target as HTMLInputElement).value;

    // In the case user enter manually a date later than today
    if (this.dayJs(date).isValid() && this.dayJs(date).isAfter(this.dayJs())) {
      date = new Date().toISOString().split('T')[0];
    }

    const filterName = `${strings.to} ${this.dayJs(date).format('ll')}`;
    this.onItemUpdated(
      {
        key: DateFilterKeys.To,
        name: filterName,
        value: date ? new Date(date).toISOString() : null,
        operator: FilterComparisonOperator.Lt
      },
      !!date // No value = unselected
    );

    if (!date && this.submittedFilterValues.length > 0) {
      this.applyFilters();
    }
  }

  private applyDateFilters() {
    // Don't apply filters if no value is selected
    if (this.toDate || this.fromDate) {
      // Keep only static filter values 'From' and 'To' and unselect the others if any
      this.selectedValues
        .filter(v => v.key !== DateFilterKeys.To && v.key !== DateFilterKeys.From)
        .forEach(v => {
          this.onFilterUpdated(this.filter.filterName, v, false);
        });

      this.selectedValues = this.selectedValues.filter(
        v => v.key === DateFilterKeys.To || v.key === DateFilterKeys.From
      );

      this.applyFilters();
    }
  }

  private getDateValue(dateKey: DateFilterKeys) {
    const date = this.selectedValues.filter(v => v.key === dateKey)[0];
    if (date) {
      return date.value.split('T')[0];
    }

    return '';
  }

  private getAllIntervals() {
    return {
      [DateFilterInterval.AnyTime]: strings.anyTime,
      [DateFilterInterval.Today]: strings.today,
      [DateFilterInterval.Past24]: strings.past24,
      [DateFilterInterval.PastWeek]: strings.pastWeek,
      [DateFilterInterval.PastMonth]: strings.pastMonth,
      [DateFilterInterval.Past3Months]: strings.past3Months,
      [DateFilterInterval.PastYear]: strings.pastYear,
      [DateFilterInterval.OlderThanAYear]: strings.olderThanAYear
    };
  }
}
