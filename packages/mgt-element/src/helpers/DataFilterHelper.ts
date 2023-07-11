import isEmpty from 'lodash-es/isEmpty';
import {
  IDataFilter,
  FilterConditionOperator,
  IDataFilterValue,
  FilterComparisonOperator
} from '../models/IDataFilter';

export class DataFilterHelper {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  static dayJs: any;

  /**
   * Build the refinement condition in FQL format
   * @param selectedFilters The selected filter array
   * @param dayJs The dayJs instance to resolve dates
   * @param encodeTokens If true, encodes the taxonomy refinement tokens in UTF-8 to work with GET requests. Javascript encodes natively in UTF-16 by default.
   */
  public static buildFqlRefinementString(selectedFilters: IDataFilter[], encodeTokens?: boolean): string[] {
    const refinementQueryConditions: string[] = [];

    selectedFilters.forEach(filter => {
      // Default operator is OR if not provided
      const operator: FilterConditionOperator = filter.operator ? filter.operator : FilterConditionOperator.OR;

      // Mutli values
      if (filter.values.length > 1) {
        let startDate: string = null;
        let endDate: string = null;

        // A refiner can have multiple values selected in a multi or mon multi selection scenario
        // The correct operator is determined by the refiner display template according to its behavior
        const conditions = filter.values
          .map((filterValue: IDataFilterValue) => {
            let value = filterValue.value;

            if (this.dayJs(value, this.dayJs.ISO_8601, true).isValid()) {
              if (
                !startDate &&
                (filterValue.operator === FilterComparisonOperator.Geq ||
                  filterValue.operator === FilterComparisonOperator.Gt)
              ) {
                startDate = value;
              }

              if (
                !endDate &&
                (filterValue.operator === FilterComparisonOperator.Lt ||
                  filterValue.operator === FilterComparisonOperator.Leq)
              ) {
                endDate = value;
              }
            }

            // If the value is null or undefined, we replace it by the FQL expression string('')
            // Otherwise the query syntax won't be vaild resuting of to an HTTP 500
            if (isEmpty(value)) {
              value = "string('')";
            }

            // Enclose the expression with quotes if the value contains spaces
            if (/\s/.test(value) && value.indexOf('range') === -1) {
              value = `"${value}"`;
            }

            return /ǂǂ/.test(value) && encodeTokens ? encodeURIComponent(value) : value;
          })
          .filter(c => c);

        if (startDate && endDate) {
          refinementQueryConditions.push(`${filter.filterName}:range(${startDate},${endDate})`);
        } else {
          refinementQueryConditions.push(`${filter.filterName}:${operator}(${conditions.join(',')})`);
        }
      } else {
        // Single value
        if (filter.values.length === 1) {
          const filterValue = filter.values[0];

          // See https://sharepoint.stackexchange.com/questions/258081/how-to-hex-encode-refiners/258161
          let refinementToken =
            /ǂǂ/.test(filterValue.value) && encodeTokens ? encodeURIComponent(filterValue.value) : filterValue.value;

          // https://docs.microsoft.com/en-us/sharepoint/dev/general-development/fast-query-language-fql-syntax-reference#fql_range_operator
          if (this.dayJs(refinementToken, this.dayJs.ISO_8601, true).isValid()) {
            if (
              filterValue.operator === FilterComparisonOperator.Gt ||
              filterValue.operator === FilterComparisonOperator.Geq
            ) {
              refinementToken = `range(${refinementToken},max)`;
            }

            // Ex: scenario ('older than a year')
            if (
              filterValue.operator === FilterComparisonOperator.Leq ||
              filterValue.operator === FilterComparisonOperator.Lt
            ) {
              refinementToken = `range(min,${refinementToken})`;
            }
          }

          // If the value is null or undefined, we replace it by the FQL expression string('')
          // Otherwise the query syntax won't be vaild resuting of to an HTTP 500
          if (isEmpty(refinementToken)) {
            refinementToken = "string('')";
          }

          // Enclose the expression with quotes if the value contains spaces
          if (/\s/.test(refinementToken) && refinementToken.indexOf('range') === -1) {
            refinementToken = `"${refinementToken}"`;
          }

          refinementQueryConditions.push(`${filter.filterName}:${refinementToken}`);
        }
      }
    });

    return refinementQueryConditions;
  }
}
