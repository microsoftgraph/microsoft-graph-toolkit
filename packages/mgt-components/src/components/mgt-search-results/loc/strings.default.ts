import { html } from 'lit-html';
import { unsafeHTML } from 'lit-html/directives/unsafe-html';

export const MgtSearchResultsStrings = {
  seeAllLink: 'See all',
  results: 'results'
};

export const MgtPaginationStrings = {
  nextBtn: 'Next',
  previousBtn: 'Previous',
  tooManyPages: 'Too many pages!',
  screenTipContent: () => html`
        <div slot="title">It seems your search returned a lot of pages!</div>
        <p>Try to narrow down your scope by specifying more precise keywords🙏</p>
    `
};

export const MgtFilterDateStrings = {
  anyTime: 'Any time',
  today: 'Today',
  past24: 'Past 24 hours',
  pastWeek: 'Past week',
  pastMonth: 'Past month',
  past3Months: 'Past 3 months',
  pastYear: 'Past year',
  olderThanAYear: 'Older than a year',
  reset: 'Reset',
  from: 'From',
  to: 'To',
  applyDates: 'Apply dates',
  selections: 'selection(s)'
};

export const MgtFilterCheckboxStrings = {
  reset: 'Reset',
  searchPlaceholder: 'Search for values...',
  apply: 'Apply',
  cancel: 'Cancel',
  selections: 'selection(s)'
};

export const MgtSearchFiltersStrings = {
  resetAllFilters: 'Reset filters',
  noFilters: 'No filter to display'
};

export const MgtSearchSortStrings = {
  sortedByRelevance: 'Sorted by relevance',
  sortDefault: 'Relevance',
  sortAscending: 'Sort ascending',
  sortDescending: 'Sort descending'
};

export const MgtSearchInputStrings = {
  searchPlaceholder: 'Search for values...',
  clearSearch: 'Clear searchbox',
  previousSearches: 'Previous searches'
};

export const MgtLanguageProviderStrings = {};

export const MgtSearchInfosStrings = {
  searchQueryResultText: (keywords): string => `Here's what we found for "${keywords}"`,
  resultCountText: (count): string => `Found ${count} results`,
  notFoundSuggestions: keywords => html`
                    <h2 class="font-mgt text-3xl mb-4">Your search for "${keywords}" did not match any content.</h2>
                    <p>Some suggestions:</p>
                    <ul class="list-disc pl-8">
                        <li>Make sure all words are spelled correctly</li>
                        <li>Try entering different keywords, more general keywords or less keywords</li>
                        <li>Maybe what you were looking for is not in the scope of the Enterprise Search? Check all the sources we index in our <a href="https://mgtaad.sharepoint.com/sites/HowToMgt/SitePages/HOW-TO-USE-ENTERPRISE-SEARCH.aspx" target="_blank" class="text-primary font-bold hover:underline">FAQ page</a>.</li>
                        <li>... Or ask for help by submitting a <a href="https://mgt.service-now.com/sos/new_ticket.do&quest;sysparm_application_id=7da6d9e4db26aa0cb1b373921f961980&amp;sysparm_action_id=a0ca11a4db66aa0cb1b373921f9619df" target="_blank" class="text-primary font-bold hover:underline">SOS ticket</a> and our colleagues will try their best to help you.</li>
                    </ul>
    `,
  didYouMean: (handlerFunction, updatedQueryString) => html`
        <p>Did you mean: "<a href="#" @click=${handlerFunction}>${unsafeHTML(updatedQueryString)}"?</a></p> 
    `
};

export const MgtErrorMessageStrings = {
  errorMessage: 'Error'
};

export const MgtExportProfilesStrings = {
  exportBtn: 'Export results',
  exportBtnLoading: 'Exporting'
};

export const strings = {
  language: 'en-us',
  _components: {
    'mgt-search-results': { ...MgtSearchResultsStrings },
    'mgt-pagination': { ...MgtPaginationStrings },
    'mgt-filter-date': { ...MgtFilterDateStrings },
    'mgt-filter-checkbox': { ...MgtFilterCheckboxStrings },
    'mgt-search-filters': { ...MgtSearchFiltersStrings },
    'mgt-search-input': { ...MgtSearchInputStrings },
    'mgt-input-autocomplete': { ...MgtSearchInputStrings },
    'mgt-language-provider': { ...MgtLanguageProviderStrings },
    'mgt-search-infos': { ...MgtSearchInfosStrings },
    'mgt-error-message': { ...MgtErrorMessageStrings },
    'mgt-sort': { ...MgtSearchSortStrings }
  }
};
