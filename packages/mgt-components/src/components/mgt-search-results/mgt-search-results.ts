/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html, HTMLTemplateResult, nothing, TemplateResult } from 'lit';
import { property, state } from 'lit/decorators.js';
import {
  CacheService,
  CacheStore,
  equals,
  prepScopes,
  Providers,
  ProviderState,
  mgtHtml,
  BetaGraph,
  BatchResponse,
  CollectionResponse,
  MgtTemplatedTaskComponent
} from '@microsoft/mgt-element';

import { schemas } from '../../graph/cacheStores';
import { strings } from './strings';
import { styles } from './mgt-search-results-css';
import {
  DirectoryObject,
  Drive,
  DriveItem,
  EntityType,
  List,
  ListItem,
  Message,
  SearchHit,
  SearchHitsContainer,
  SearchRequest,
  SearchResponse,
  Site
} from '@microsoft/microsoft-graph-types';
import { SearchRequest as BetaSearchRequest } from '@microsoft/microsoft-graph-types-beta';
import {
  getIsResponseCacheEnabled,
  getNameFromUrl,
  getRelativeDisplayDate,
  getResponseInvalidationTime,
  sanitizeSummary,
  trimFileExtension
} from '../../utils/Utils';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';
import { fluentSkeleton, fluentButton, fluentTooltip, fluentDivider } from '@fluentui/web-components';
import { registerFluentComponents } from '../../utils/FluentComponents';
import { CacheResponse } from '../CacheResponse';
import { registerComponent } from '@microsoft/mgt-element';
import { registerMgtFileComponent } from '../mgt-file/mgt-file';
import { registerMgtPersonComponent } from '../mgt-person/mgt-person';

/**
 * Object representing a thumbnail
 */
interface Thumbnail {
  /**
   * The url of the Thumbnail
   */
  url?: string;
}

/**
 * Object representing a Binary Thumbnail
 */
interface BinaryThumbnail {
  /**
   * The url of the Thumbnail
   */
  url?: string;

  /**
   * The web Url of the Thumbnail
   */
  thumbnailWebUrl?: string;
}

/**
 * Object representing a Search Answer
 */
interface Answer {
  '@odata.type': string;
  displayName?: string;
  description?: string;
  webUrl?: string;
}

/**
 * Object representing a search resource supporting thumbnails
 */
interface ThumbnailResource {
  thumbnail: Thumbnail;
}

interface UserResource {
  lastModifiedBy?: {
    user?: {
      email?: string;
    };
  };
  userPrincipalName?: string;
}

/**
 * Object representing a Search Resource
 */
type SearchResource = Partial<
  DriveItem & Site & List & Message & ListItem & Drive & DirectoryObject & Answer & ThumbnailResource & UserResource
>;

/**
 * Object representing a full Search Response
 */
type SearchResponseCollection = CollectionResponse<SearchResponse>;

export const registerMgtSearchResultsComponent = () => {
  registerFluentComponents(fluentSkeleton, fluentButton, fluentTooltip, fluentDivider);

  registerMgtFileComponent();
  registerMgtPersonComponent();
  registerComponent('search-results', MgtSearchResults);
};

/**
 * **Preview component** Custom element for making Microsoft Graph get queries.
 * Component may change before general availability release.
 *
 * @fires {CustomEvent<undefined>} updated - Fired when the component is updated
 * @fires {CustomEvent<DataChangedDetail>} dataChange - Fired when data changes
 *
 * @cssprop --answer-border-radius - {Length} Border radius of an answer
 * @cssprop --answer-box-shadow - {Length} Box shadow of an answer
 * @cssprop --answer-border - {Length} Border of an answer
 * @cssprop --answer-padding - {Length} Padding of an answer
 *
 * @class mgt-search-results
 * @extends {MgtTemplatedTaskComponent}
 */
export class MgtSearchResults extends MgtTemplatedTaskComponent {
  /**
   * Default page size is 10
   */
  private _size = 10;

  /**
   * Array of styles to apply to the element. The styles should be defined
   * user the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * Gets all the localization strings for the component
   */
  protected get strings() {
    return strings;
  }

  private _queryString: string;

  /**
   * The query to send to Microsoft Search
   *
   * @type {string}
   * @memberof MgtSearchResults
   */
  public get queryString(): string {
    return this._queryString;
  }
  @property({
    attribute: 'query-string',
    type: String
  })
  public set queryString(value: string) {
    if (this._queryString !== value) {
      this._queryString = value;
      this.currentPage = 1;
    }
  }

  /**
   * Query template to use in complex search scenarios
   * Query Templates are currently supported only on the beta endpoint
   */
  @property({
    attribute: 'query-template',
    type: String
  })
  public queryTemplate: string;

  /**
   * One or more types of resources expected in the response.
   * Possible values are: list, site, listItem, message, event,
   * drive, driveItem, externalItem.
   *
   * @type {string[]}
   * @memberof MgtSearchResults
   */
  @property({
    attribute: 'entity-types',
    converter: value => {
      return value.split(',').map(v => v.trim());
    },
    type: String
  })
  public entityTypes: string[] = ['driveItem', 'listItem', 'site'];

  /**
   * The scopes to request
   *
   * @type {string[]}
   * @memberof MgtSearchResults
   */
  @property({
    attribute: 'scopes',
    converter: (value, _type) => {
      return value ? value.toLowerCase().split(',') : null;
    }
  })
  public scopes: string[] = [];

  /**
   * Content sources to use with External Items
   *
   * @type {string[]}
   * @memberof MgtSearchResults
   */
  @property({
    attribute: 'content-sources',
    converter: (value, _type) => {
      return value ? value.toLowerCase().split(',') : null;
    }
  })
  public contentSources: string[] = [];

  /**
   * Api version to use for request
   *
   * @type {string}
   * @memberof MgtSearchResults
   */
  @property({
    attribute: 'version',
    reflect: true,
    type: String
  })
  public version = 'v1.0';

  /**
   * Specifies the offset for the search results.
   * Offset 0 returns the very first result.
   *
   * @type {number}
   * @memberof MgtSearchResults
   */
  public get from(): number {
    return (this.currentPage - 1) * this.size;
  }

  /**
   * The size of the page to be retrieved.
   * The maximum value is 1000.
   *
   * @type {number}
   * @memberof MgtSearchResults
   */
  public get size(): number {
    return this._size;
  }
  @property({
    attribute: 'size',
    reflect: true,
    type: Number
  })
  public set size(value) {
    if (value > this.maxPageSize) {
      this._size = this.maxPageSize;
    } else {
      this._size = value;
    }
  }

  /**
   * The maximum number of pages to be clickable
   * in the paging control
   *
   * @type {number}
   * @memberof MgtSearchResults
   */
  @property({
    attribute: 'paging-max',
    reflect: true,
    type: Number
  })
  public pagingMax = 7;

  /**
   * Sets whether the result thumbnail should be fetched
   * from the Microsoft Graph
   *
   * @type {boolean}
   * @memberof MgtSearchResults
   */
  @property({
    attribute: 'fetch-thumbnail',
    type: Boolean
  })
  public fetchThumbnail: boolean;

  /**
   * Contains the fields to be returned for each resource
   *
   * @type {string[]}
   * @memberof MgtSearchResults
   */
  @property({
    attribute: 'fields',
    converter: value => {
      return value.split(',').map(v => v.trim());
    },
    type: String
  })
  public fields: string[];

  /**
   * This triggers hybrid sort for messages : the first 3 messages are the most relevant.
   * This property is only applicable to entityType=message
   *
   * @type {boolean}
   * @memberof MgtSearchResults
   */
  @property({
    attribute: 'enable-top-results',
    reflect: true,
    type: Boolean
  })
  public enableTopResults = false;

  /**
   * Enables cache on the response from the specified resource
   * default = false
   *
   * @type {boolean}
   * @memberof MgtSearchResults
   */
  @property({
    attribute: 'cache-enabled',
    reflect: true,
    type: Boolean
  })
  public cacheEnabled = false;

  /**
   * Invalidation period of the cache for the responses in milliseconds
   *
   * @type {number}
   * @memberof MgtSearchResults
   */
  @property({
    attribute: 'cache-invalidation-period',
    reflect: true,
    type: Number
  })
  public cacheInvalidationPeriod = 30000;

  /**
   * Gets or sets the response of the request
   *
   * @type any
   * @memberof MgtSearchResults
   */
  @state() private response: SearchResponseCollection;

  private isRefreshing = false;
  private get searchEndpoint() {
    return '/search/query';
  }
  private get maxPageSize() {
    return 1000;
  }
  private readonly defaultFields: string[] = [
    'webUrl',
    'lastModifiedBy',
    'lastModifiedDateTime',
    'summary',
    'displayName',
    'name'
  ];

  @property({ attribute: false })
  public currentPage = 1;

  constructor() {
    super();
  }

  protected args(): unknown[] {
    return [
      this.providerState,
      this.queryString,
      this.queryTemplate,
      this.entityTypes,
      this.contentSources,
      this.scopes,
      this.version,
      this.size,
      this.fetchThumbnail,
      this.fields,
      this.enableTopResults,
      this.currentPage
    ];
  }

  /**
   * Refresh the data
   *
   * @param {boolean} [hardRefresh=false]
   * if false (default), the component will only update if the data changed
   * if true, the data will be first cleared and reloaded completely
   * @memberof MgtSearchResults
   */
  public refresh(hardRefresh = false) {
    this.isRefreshing = true;
    if (hardRefresh) {
      this.clearState();
    }
    void this._task.run();
  }

  /**
   * Clears state of the component
   *
   * @protected
   * @memberof MgtSearchResults
   */
  protected clearState(): void {
    this.response = null;
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected renderContent = (): TemplateResult => {
    let renderedTemplate: TemplateResult = null;
    let headerTemplate: TemplateResult = null;
    let footerTemplate: TemplateResult = null;

    if (this.hasTemplate('header')) {
      headerTemplate = this.renderTemplate('header', this.response);
    }

    footerTemplate = this.renderFooter(this.response?.value[0]?.hitsContainers[0]);

    if (this.response && this.hasTemplate('default')) {
      renderedTemplate = this.renderTemplate('default', this.response) || html``;
    } else if (this.response?.value[0]?.hitsContainers[0]) {
      renderedTemplate = html`${this.response?.value[0]?.hitsContainers[0]?.hits?.map(result =>
        this.renderResult(result)
      )}`;
    } else if (this.hasTemplate('no-data')) {
      renderedTemplate = this.renderTemplate('no-data', null);
    } else {
      renderedTemplate = html``;
    }

    return html`
      ${headerTemplate}
      <div class="search-results">
        ${renderedTemplate}
      </div>
      ${footerTemplate}`;
  };

  /**
   * load state into the component.
   *
   * @protected
   * @returns
   * @memberof MgtSearchResults
   */
  protected async loadState() {
    const provider = Providers.globalProvider;

    this.error = null;

    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    if (this.queryString) {
      try {
        const requestOptions = this.getRequestOptions();

        let cache: CacheStore<CacheResponse>;
        const key = JSON.stringify({
          endpoint: `${this.version}${this.searchEndpoint}`,
          requestOptions
        });
        let response: SearchResponseCollection = null;

        if (this.shouldRetrieveCache()) {
          cache = CacheService.getCache<CacheResponse>(schemas.search, schemas.search.stores.responses);
          const result: CacheResponse = getIsResponseCacheEnabled() ? await cache.getValue(key) : null;
          if (result && getResponseInvalidationTime(this.cacheInvalidationPeriod) > Date.now() - result.timeCached) {
            response = JSON.parse(result.response) as SearchResponseCollection;
          }
        }

        if (!response) {
          const graph = provider.graph.forComponent(this);
          let request = graph.api(this.searchEndpoint).version(this.version);

          if (this.scopes?.length) {
            request = request.middlewareOptions(prepScopes(this.scopes));
          }

          response = (await request.post({ requests: [requestOptions] })) as SearchResponseCollection;

          if (this.fetchThumbnail) {
            const thumbnailBatch = graph.createBatch<BinaryThumbnail>();
            const thumbnailBatchBeta = BetaGraph.fromGraph(graph).createBatch<BinaryThumbnail>();

            const hits =
              response.value?.length && response.value[0].hitsContainers?.length
                ? response.value[0].hitsContainers[0]?.hits ?? []
                : [];
            for (const element of hits) {
              const resource = element.resource as SearchResource;
              if (
                (resource.size > 0 || resource.webUrl?.endsWith('.aspx')) &&
                (resource['@odata.type'] === '#microsoft.graph.driveItem' ||
                  resource['@odata.type'] === '#microsoft.graph.listItem')
              ) {
                if (resource['@odata.type'] === '#microsoft.graph.listItem') {
                  thumbnailBatchBeta.get(
                    element.hitId.toString(),
                    `/sites/${resource.parentReference.siteId}/pages/${resource.id}`
                  );
                } else {
                  thumbnailBatch.get(
                    element.hitId.toString(),
                    `/drives/${resource.parentReference.driveId}/items/${resource.id}/thumbnails/0/medium`
                  );
                }
              }
            }

            /**
             * Based on the batch response, augment the search result resource with the thumbnail url
             *
             * @param thumbnailResponse
             */
            const augmentResponse = (thumbnailResponse: Map<string, BatchResponse<BinaryThumbnail>>) => {
              if (thumbnailResponse && thumbnailResponse.size > 0) {
                for (const [k, value] of thumbnailResponse) {
                  const result: SearchHit = response.value[0].hitsContainers[0].hits[k] as SearchHit;
                  const thumbnail: Thumbnail =
                    result.resource['@odata.type'] === '#microsoft.graph.listItem'
                      ? { url: value.content.thumbnailWebUrl }
                      : { url: value.content.url };
                  (result.resource as SearchResource).thumbnail = thumbnail;
                }
              }
            };

            try {
              augmentResponse(await thumbnailBatch.executeAll());
              augmentResponse(await thumbnailBatchBeta.executeAll());
            } catch {
              // no-op
            }
          }

          if (this.shouldUpdateCache() && response) {
            cache = CacheService.getCache<CacheResponse>(schemas.search, schemas.search.stores.responses);
            await cache.putValue(key, { response: JSON.stringify(response) });
          }
        }

        if (!equals(this.response, response)) {
          this.response = response;
        }
      } catch (e: unknown) {
        this.error = e as Error;
      }

      if (this.response) {
        this.error = null;
      }
    } else {
      this.response = null;
    }
    this.isRefreshing = false;
    this.fireCustomEvent('dataChange', { response: this.response, error: this.error as Error });
  }

  /**
   * Render the loading state.
   *
   * @protected
   * @returns
   * @memberof MgtSearchResults
   */
  protected readonly renderLoading = (): TemplateResult => {
    return (
      this.renderTemplate('loading', null) ||
      // creates an array of n items where n is the current max number of results, this builds a shimmer for that many results
      html`
        ${[...Array<number>(this.size)].map(() => {
          return html`
            <div class="search-result">
              <div class="search-result-grid">
                <div class="search-result-icon">
                  <fluent-skeleton class="search-result-icon__shimmer" shape="rect" shimmer></fluent-skeleton>
                </div>
                <div class="searc-result-content">
                  <div class="search-result-name">
                    <fluent-skeleton class="search-result-name__shimmer" shape="rect" shimmer></fluent-skeleton>
                  </div>
                  <div class="search-result-info">
                    <div class="search-result-author">
                      <fluent-skeleton class="search-result-author__shimmer" shape="circle" shimmer></fluent-skeleton>
                    </div>
                    <div class="search-result-date">
                      <fluent-skeleton class="search-result-date__shimmer" shape="rect" shimmer></fluent-skeleton>
                    </div>
                  </div>
                  <fluent-skeleton class="search-result-content__shimmer" shape="rect" shimmer></fluent-skeleton>
                  <fluent-skeleton class="search-result-content__shimmer" shape="rect" shimmer></fluent-skeleton>
                </div>
                ${
                  this.fetchThumbnail &&
                  html`
                    <div class="search-result-thumbnail">
                      <fluent-skeleton class="search-result-thumbnail__shimmer" shape="rect" shimmer></fluent-skeleton>
                    </div>
                  `
                }
              </div>
              <fluent-divider></fluent-divider>
            </div>
          `;
        })}
       `
    );
  };

  /**
   * Render the result item.
   *
   * @protected
   * @returns
   * @memberof MgtSearchResults
   */
  protected renderResult(result: SearchHit): TemplateResult {
    const type = this.getResourceType(result.resource);
    if (this.hasTemplate(`result-${type}`)) {
      return this.renderTemplate(`result-${type}`, result, result.hitId);
    } else {
      switch (result.resource['@odata.type']) {
        case '#microsoft.graph.driveItem':
          return this.renderDriveItem(result);
        case '#microsoft.graph.site':
          return this.renderSite(result);
        case '#microsoft.graph.person':
          return this.renderPerson(result);
        case '#microsoft.graph.drive':
        case '#microsoft.graph.list':
          return this.renderList(result);
        case '#microsoft.graph.listItem':
          return this.renderListItem(result);
        case '#microsoft.graph.search.bookmark':
          return this.renderBookmark(result);
        case '#microsoft.graph.search.acronym':
          return this.renderAcronym(result);
        case '#microsoft.graph.search.qna':
          return this.renderQnA(result);
        default:
          return this.renderDefault(result);
      }
    }
  }

  /**
   * Renders the footer with pages if required
   *
   * @param hitsContainer Search results
   */
  private renderFooter(hitsContainer: SearchHitsContainer) {
    if (this.pagingRequired(hitsContainer)) {
      const pages = this.getActivePages(hitsContainer.total);

      return html`
        <div class="search-results-pages">
          ${this.renderPreviousPage()}
          ${this.renderFirstPage(pages)}
          ${this.renderAllPages(pages)}
          ${this.renderNextPage()}
        </div>
      `;
    }
  }

  /**
   * Validates if paging is required based on the provided results
   *
   * @param hitsContainer
   */
  private pagingRequired(hitsContainer: SearchHitsContainer) {
    return hitsContainer?.moreResultsAvailable || this.currentPage * this.size < hitsContainer?.total;
  }

  /**
   * Gets a list of active pages to render for paging purposes
   *
   * @param totalResults Total number of results of the search query
   */
  private getActivePages(totalResults: number) {
    const getFirstPage = () => {
      const medianPage = this.currentPage - Math.floor(this.pagingMax / 2) - 1;

      if (medianPage >= Math.floor(this.pagingMax / 2)) {
        return medianPage;
      } else {
        return 0;
      }
    };

    const pages: number[] = [];
    const firstPage = getFirstPage();

    if (firstPage + 1 > this.pagingMax - this.currentPage || this.pagingMax === this.currentPage) {
      for (
        let i = firstPage + 1;
        i < Math.ceil(totalResults / this.size) &&
        i < this.pagingMax + (this.currentPage - 1) &&
        pages.length < this.pagingMax - 2;
        ++i
      ) {
        pages.push(i + 1);
      }
    } else {
      for (let i = firstPage; i < this.pagingMax; ++i) {
        pages.push(i + 1);
      }
    }

    return pages;
  }

  /**
   * Renders all sequential pages buttons
   *
   * @param pages
   */
  private renderAllPages(pages: number[]) {
    return html`
      ${pages.map(
        page =>
          html`
            <fluent-button
              title="${strings.page} ${page}"
              appearance="stealth"
              class="${page === this.currentPage ? 'search-results-page-active' : 'search-results-page'}"
              @click="${() => this.onPageClick(page)}">
                ${page}
            </fluent-button>`
      )}`;
  }

  /**
   * Renders the "First page" button
   *
   * @param pages
   */
  private renderFirstPage(pages: number[]) {
    return html`
      ${
        pages.some(page => page === 1)
          ? nothing
          : html`
              <fluent-button
                 title="${strings.page} 1"
                 appearance="stealth"
                 class="search-results-page"
                 @click="${this.onFirstPageClick}">
                 1
               </fluent-button>`
            ? html`
              <fluent-button
                id="page-back-dot"
                appearance="stealth"
                class="search-results-page"
                title="${this.getDotButtonTitle()}"
                @click="${() => this.onPageClick(this.currentPage - Math.ceil(this.pagingMax / 2))}"
              >
                ...
              </fluent-button>`
            : nothing
      }`;
  }

  /**
   * Constructs the "dot dot dot" button title
   */
  private getDotButtonTitle() {
    return `${strings.back} ${Math.ceil(this.pagingMax / 2)} ${strings.pages}`;
  }

  /**
   * Renders the "Previous page" button
   */
  private renderPreviousPage() {
    return this.currentPage > 1
      ? html`
          <fluent-button
            appearance="stealth"
            class="search-results-page"
            title="${strings.back}"
            @click="${this.onPageBackClick}">
              ${getSvg(SvgIcon.ChevronLeft)}
            </fluent-button>`
      : nothing;
  }

  /**
   * Renders the "Next page" button
   */
  private renderNextPage() {
    return !this.isLastPage()
      ? html`
          <fluent-button
            appearance="stealth"
            class="search-results-page"
            title="${strings.next}"
            aria-label="${strings.next}"
            @click="${this.onPageNextClick}">
              ${getSvg(SvgIcon.ChevronRight)}
            </fluent-button>`
      : nothing;
  }

  /**
   * Triggers a specific page click
   *
   * @param pageNumber
   */
  private onPageClick(pageNumber: number) {
    this.currentPage = pageNumber;
    this.scrollToFirstResult();
  }

  /**
   * Triggers a first page click
   *
   */
  private readonly onFirstPageClick = () => {
    this.currentPage = 1;
    this.scrollToFirstResult();
  };

  /**
   * Triggers a previous page click
   */
  private readonly onPageBackClick = () => {
    this.currentPage--;
    this.scrollToFirstResult();
  };

  /**
   * Triggers a next page click
   */
  private readonly onPageNextClick = () => {
    this.currentPage++;
    this.scrollToFirstResult();
  };

  /**
   * Validates if the current page is the last page of the collection
   */
  private isLastPage() {
    return this.currentPage === Math.ceil(this.response.value[0].hitsContainers[0].total / this.size);
  }

  /**
   * Scroll to the top of the search results
   */
  private scrollToFirstResult() {
    const target = this.renderRoot.querySelector('.search-results');
    target.scrollIntoView({
      block: 'start',
      behavior: 'smooth'
    });
  }

  /**
   * Gets the resource type (entity) of a search result
   *
   * @param resource
   */
  private getResourceType(resource: SearchResource) {
    return resource['@odata.type'].split('.').pop();
  }

  /**
   * Renders a driveItem entity
   *
   * @param result
   */
  private renderDriveItem(result: SearchHit) {
    const resource = result.resource as SearchResource;
    return mgtHtml`
      <div class="search-result-grid">
        <div class="search-result-icon">
          <mgt-file
            .fileDetails="${result.resource}"
            view="image"
            class="file-icon">
          </mgt-file>
        </div>
        <div class="search-result-content">
          <div class="search-result-name">
            <a href="${resource.webUrl}?Web=1" target="_blank">${trimFileExtension(resource.name)}</a>
          </div>
          <div class="search-result-info">
            <div class="search-result-author">
              <mgt-person
                person-query=${resource.lastModifiedBy.user.email}
                view="oneline"
                person-card="hover"
                show-presence="true">
              </mgt-person>
            </div>
            <div class="search-result-date">
              &nbsp; ${strings.modified} ${getRelativeDisplayDate(new Date(resource.lastModifiedDateTime))}
            </div>
          </div>
          <div class="search-result-summary" .innerHTML="${sanitizeSummary(result.summary)}"></div>
        </div>
        ${
          resource.thumbnail?.url &&
          html`
          <div class="search-result-thumbnail">
            <a href="${resource.webUrl}" target="_blank"><img alt="${resource.name}" src="${resource.thumbnail?.url}" /></a>
          </div>`
        }

      </div>
      <fluent-divider></fluent-divider>
    `;
  }

  /**
   * Renders a site entity
   *
   * @param result
   * @returns
   */
  private renderSite(result: SearchHit): HTMLTemplateResult {
    const resource = result.resource as SearchResource;
    return html`
      <div class="search-result-grid">
        <div class="search-result-icon">
          ${this.getResourceIcon(resource)}
        </div>
        <div class="searc-result-content">
          <div class="search-result-name">
            <a href="${resource.webUrl}" target="_blank">${resource.displayName}</a>
          </div>
          <div class="search-result-info">
            <div class="search-result-url">
              <a href="${resource.webUrl}" target="_blank">${resource.webUrl}</a>
            </div>
          </div>
          <div class="search-result-summary" .innerHTML="${sanitizeSummary(result.summary)}"></div>
        </div>
      </div>
      <fluent-divider></fluent-divider>
    `;
  }

  /**
   * Renders a list entity
   *
   * @param result
   * @returns
   */
  private renderList(result: SearchHit): HTMLTemplateResult {
    const resource = result.resource as SearchResource;
    return mgtHtml`
      <div class="search-result-grid">
        <div class="search-result-icon">
          <mgt-file
            .fileDetails="${result.resource}"
            view="image">
          </mgt-file>
        </div>
        <div class="search-result-content">
          <div class="search-result-name">
            <a href="${resource.webUrl}?Web=1" target="_blank">
              ${trimFileExtension(resource.name || getNameFromUrl(resource.webUrl))}
            </a>
          </div>
          <div class="search-result-summary" .innerHTML="${sanitizeSummary(result.summary)}"></div>
        </div>
      </div>
      <fluent-divider></fluent-divider>
    `;
  }

  /**
   * Renders a listItem entity
   *
   * @param result
   * @returns
   */
  private renderListItem(result: SearchHit): HTMLTemplateResult {
    const resource = result.resource as SearchResource;
    return mgtHtml`
      <div class="search-result-grid">
        <div class="search-result-icon">
          ${resource.webUrl.endsWith('.aspx') ? getSvg(SvgIcon.News) : getSvg(SvgIcon.FileOuter)}
        </div>
        <div class="search-result-content">
          <div class="search-result-name">
            <a href="${resource.webUrl}?Web=1" target="_blank">
              ${trimFileExtension(resource.name || getNameFromUrl(resource.webUrl))}
            </a>
          </div>
          <div class="search-result-info">
            <div class="search-result-author">
              <mgt-person
                person-query=${resource.lastModifiedBy.user.email}
                view="oneline"
                person-card="hover"
                show-presence="true">
              </mgt-person>
            </div>
            <div class="search-result-date">
              &nbsp; ${strings.modified} ${getRelativeDisplayDate(new Date(resource.lastModifiedDateTime))}
            </div>
          </div>
          <div class="search-result-summary" .innerHTML="${sanitizeSummary(result.summary)}"></div>
        </div>
        ${
          resource.thumbnail?.url &&
          html`
          <div class="search-result-thumbnail">
            <a href="${resource.webUrl}" target="_blank"><img alt="${trimFileExtension(
              resource.name || getNameFromUrl(resource.webUrl)
            )}" src="${resource.thumbnail?.url || nothing}" /></a>
          </div>`
        }
      </div>
      <fluent-divider></fluent-divider>
    `;
  }

  /**
   * Renders a person entity
   *
   * @param result
   * @returns
   */
  private renderPerson(result: SearchHit): HTMLTemplateResult {
    const resource = result.resource as SearchResource;
    return mgtHtml`
      <div class="search-result">
        <mgt-person
          view="fourlines"
          person-query=${resource.userPrincipalName}
          person-card="hover"
          show-presence="true">
        </mgt-person>
      </div>
      <fluent-divider></fluent-divider>
    `;
  }

  /**
   * Renders a bookmark entity
   *
   * @param result
   */
  private renderBookmark(result: SearchHit) {
    return this.renderAnswer(result, SvgIcon.DoubleBookmark);
  }

  /**
   * Renders an acronym entity
   *
   * @param result
   */
  private renderAcronym(result: SearchHit) {
    return this.renderAnswer(result, SvgIcon.BookOpen);
  }

  /**
   * Renders a qna entity
   *
   * @param result
   */
  private renderQnA(result: SearchHit) {
    return this.renderAnswer(result, SvgIcon.BookQuestion);
  }

  /**
   * Renders an answer entity
   *
   * @param result
   */
  private renderAnswer(result: SearchHit, icon: SvgIcon) {
    const resource = result.resource as SearchResource;
    return html`
      <div class="search-result-grid search-result-answer">
        <div class="search-result-icon">
          ${getSvg(icon)}
        </div>
        <div class="search-result-content">
          <div class="search-result-name">
            <a href="${this.getResourceUrl(resource)}?Web=1" target="_blank">${resource.displayName}</a>
          </div>
          <div class="search-result-summary">${resource.description}</div>
        </div>
      </div>
      <fluent-divider></fluent-divider>
    `;
  }

  /**
   * Renders any entity
   *
   * @param result
   */
  private renderDefault(result: SearchHit) {
    const resource = result.resource as SearchResource;
    const resourceUrl = this.getResourceUrl(resource);
    return html`
      <div class="search-result-grid">
        <div class="search-result-icon">
          ${this.getResourceIcon(resource)}
        </div>
        <div class="search-result-content">
          <div class="search-result-name">
            ${
              resourceUrl
                ? html`
                  <a href="${resourceUrl}?Web=1" target="_blank">${this.getResourceName(resource)}</a>
                `
                : html`
                  ${this.getResourceName(resource)}
                `
            }
          </div>
          <div class="search-result-summary" .innerHTML="${this.getResultSummary(result)}"></div>
        </div>
      </div>
      <fluent-divider></fluent-divider>
    `;
  }

  /**
   * Gets default resource URLs
   *
   * @param resource
   */
  private getResourceUrl(resource: SearchResource) {
    return resource.webUrl || /* resource.url ||*/ resource.webLink || null;
  }

  /**
   * Gets default resource Names
   *
   * @param resource
   */
  private getResourceName(resource: SearchResource) {
    return resource.displayName || resource.subject || trimFileExtension(resource.name);
  }

  /**
   * Gets default result summary
   *
   * @param resource
   */
  private getResultSummary(result: SearchHit) {
    return sanitizeSummary(result.summary || (result.resource as SearchResource)?.description || null);
  }

  /**
   * Gets default resource icon
   *
   * @param resource
   */
  private getResourceIcon(resource: SearchResource) {
    switch (resource['@odata.type']) {
      case '#microsoft.graph.site':
        return getSvg(SvgIcon.Globe);
      case '#microsoft.graph.message':
        return getSvg(SvgIcon.Email);
      case '#microsoft.graph.event':
        return getSvg(SvgIcon.Event);
      case 'microsoft.graph.chatMessage':
        return getSvg(SvgIcon.SmallChat);
      default:
        return getSvg(SvgIcon.FileOuter);
    }
  }

  /**
   * Validates if cache should be retrieved
   *
   * @returns
   */
  private shouldRetrieveCache(): boolean {
    return getIsResponseCacheEnabled() && this.cacheEnabled && !this.isRefreshing;
  }

  /**
   * Validates if cache should be updated
   *
   * @returns
   */
  private shouldUpdateCache(): boolean {
    return getIsResponseCacheEnabled() && this.cacheEnabled;
  }

  /**
   * Builds the appropriate RequestOption for the search query
   *
   * @returns
   */
  private getRequestOptions(): SearchRequest | BetaSearchRequest {
    const requestOptions: SearchRequest = {
      entityTypes: this.entityTypes as EntityType[],
      query: {
        queryString: this.queryString
      },
      from: this.from ? this.from : undefined,
      size: this.size ? this.size : undefined,
      fields: this.getFields(),
      enableTopResults: this.enableTopResults ? this.enableTopResults : undefined
    };

    if (this.entityTypes.includes('externalItem')) {
      requestOptions.contentSources = this.contentSources;
    }

    if (this.version === 'beta') {
      (requestOptions as BetaSearchRequest).query.queryTemplate = this.queryTemplate ? this.queryTemplate : undefined;
    }

    return requestOptions;
  }

  /**
   * Gets the fields and default fields for default render methods
   *
   * @returns
   */
  private getFields(): string[] {
    if (this.fields) {
      return this.defaultFields.concat(this.fields);
    }

    return undefined;
  }
}
