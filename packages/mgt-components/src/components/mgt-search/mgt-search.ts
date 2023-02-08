/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { property } from 'lit/decorators.js';
import {
  CacheItem,
  CacheService,
  CacheStore,
  equals,
  MgtTemplatedComponent,
  prepScopes,
  Providers,
  ProviderState,
  customElement
} from '@microsoft/mgt-element';

import { schemas } from '../../graph/cacheStores';

/**
 * Object to be stored in cache representing a generic query
 */
interface CacheResponse extends CacheItem {
  /**
   * json representing a response as string
   */
  response?: string;
}

/**
 * Defines the expiration time
 */
const getResponseInvalidationTime = (currentInvalidationPeriod: number): number =>
  currentInvalidationPeriod ||
  CacheService.config.response.invalidationPeriod ||
  CacheService.config.defaultInvalidationPeriod;

/**
 * Whether the response store is enabled
 */
const getIsResponseCacheEnabled = (): boolean =>
  CacheService.config.response.isEnabled && CacheService.config.isEnabled;

/**
 * Custom element for making Microsoft Graph get queries
 *
 * @fires {CustomEvent<DataChangedDetail>} dataChange - Fired when data changes
 *
 * @export
 * @class mgt-get
 * @extends {MgtTemplatedComponent}
 */
@customElement('search')
// @customElement('mgt-search')
export class MgtSearch extends MgtTemplatedComponent {
  /**
   * The query to send to Microsoft Search
   *
   * @type {string}
   * @memberof MgtSearch
   */
  @property({
    attribute: 'query-string',
    reflect: true,
    type: String
  })
  public queryString: string;

  /**
   * One or more types of resources expected in the response.
   * Possible values are: list, site, listItem, message, event,
   * drive, driveItem, externalItem.
   *
   * @type {string}
   * @memberof MgtSearch
   */
  @property({
    attribute: 'entity-types',
    converter: value => {
      return value.split(',').map(v => v.trim());
    },
    type: String
  })
  public entityTypes: string[];

  /**
   * The scopes to request
   *
   * @type {string[]}
   * @memberof MgtSearch
   */
  @property({
    attribute: 'scopes',
    converter: (value, type) => {
      return value ? value.toLowerCase().split(',') : null;
    },
    reflect: true
  })
  public scopes: string[] = [];

  /**
   * Api version to use for request
   *
   * @type {string}
   * @memberof MgtSearch
   */
  @property({
    attribute: 'version',
    reflect: true,
    type: String
  })
  public version: string = 'v1.0';

  /**
   * Specifies the offset for the search results.
   * Offset 0 returns the very first result.
   *
   * @type {number}
   * @memberof MgtSearch
   */
  @property({
    attribute: 'from',
    reflect: true,
    type: Number
  })
  public from: number = 0;

  /**
   * The size of the page to be retrieved.
   * The maximum value is 1000.
   *
   * @type {number}
   * @memberof MgtSearch
   */
  @property({
    attribute: 'size',
    reflect: true,
    type: Number
  })
  public size: number = 10;

  /**
   * Contains the fields to be returned for each resource
   *
   * @type {string[]}
   * @memberof MgtSearch
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
   * @memberof MgtSearch
   */
  @property({
    attribute: 'enable-top-results',
    reflect: true,
    type: Boolean
  })
  public enableTopResults: boolean = false;

  /**
   * Enables cache on the response from the specified resource
   * default = false
   *
   * @type {boolean}
   * @memberof MgtSearch
   */
  @property({
    attribute: 'cache-enabled',
    reflect: true,
    type: Boolean
  })
  public cacheEnabled: boolean = false;

  /**
   * Invalidation period of the cache for the responses in milliseconds
   *
   * @type {number}
   * @memberof MgtSearch
   */
  @property({
    attribute: 'cache-invalidation-period',
    reflect: true,
    type: Number
  })
  public cacheInvalidationPeriod: number = 0;

  /**
   * Gets or sets the response of the request
   *
   * @type any
   * @memberof MgtSearch
   */
  @property({ attribute: false }) public response: any;

  /**
   *
   * Gets or sets the error (if any) of the request
   * @type any
   * @memberof MgtSearch
   */
  @property({ attribute: false }) public error: any;

  private isRefreshing: boolean = false;
  private readonly SEARCH_ENDPOINT: string = '/search/query';

  /**
   * Synchronizes property values when attributes change.
   *
   * @param {*} name
   * @param {*} oldValue
   * @param {*} newValue
   * @memberof MgtSearch
   */
  public attributeChangedCallback(name, oldval, newval) {
    super.attributeChangedCallback(name, oldval, newval);
    this.requestStateUpdate();
  }

  /**
   * Refresh the data
   *
   * @param {boolean} [hardRefresh=false]
   * if false (default), the component will only update if the data changed
   * if true, the data will be first cleared and reloaded completely
   * @memberof MgtSearch
   */
  public refresh(hardRefresh = false) {
    this.isRefreshing = true;
    if (hardRefresh) {
      this.clearState();
    }
    this.requestStateUpdate(hardRefresh);
  }

  /**
   * Clears state of the component
   *
   * @protected
   * @memberof MgtSearch
   */
  protected clearState(): void {
    this.response = null;
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    let renderedTemplate = null;
    let headerTemplate = null;
    let footerTemplate = null;

    // tslint:disable-next-line: no-string-literal
    if (this.hasTemplate('header')) {
      headerTemplate = this.renderTemplate('header', this.response);
    }

    if (this.hasTemplate('footer')) {
      footerTemplate = this.renderTemplate('footer', this.response);
    }

    if (this.isLoadingState) {
      renderedTemplate = this.renderTemplate('loading', null);
    } else if (this.error) {
      renderedTemplate = this.renderTemplate('error', this.error);
      // tslint:disable-next-line: no-string-literal
    } else if (this.hasTemplate('result') && this.response && this.response?.value[0]?.hitsContainers[0]) {
      let valueContent;

      if (Array.isArray(this.response.value[0].hitsContainers[0].hits)) {
        let loading = null;
        if (this.isLoadingState) {
          loading = this.renderTemplate('loading', null);
        }
        valueContent = html`
          ${this.response?.value[0]?.hitsContainers[0]?.hits.map(v =>
            this.renderTemplate('result', v, v.hitId)
          )} ${loading}
        `;
      }

      if (this.hasTemplate('default')) {
        const defaultContent = this.renderTemplate('default', this.response);

        // tslint:disable-next-line: no-string-literal
        if ((this.templates['result'] as any).templateOrder > (this.templates['default'] as any).templateOrder) {
          renderedTemplate = html`
            ${defaultContent}${valueContent}
          `;
        } else {
          renderedTemplate = html`
            ${valueContent}${defaultContent}
          `;
        }
      } else {
        renderedTemplate = valueContent;
      }
    } else if (this.response) {
      renderedTemplate = this.renderTemplate('default', this.response) || html``;
    } else if (this.hasTemplate('no-data')) {
      renderedTemplate = this.renderTemplate('no-data', null);
    } else {
      renderedTemplate = html``;
    }

    return html`${headerTemplate}${renderedTemplate}${footerTemplate}`;
  }

  /**
   * load state into the component.
   *
   * @protected
   * @returns
   * @memberof MgtSearch
   */
  protected async loadState() {
    const provider = Providers.globalProvider;

    this.error = null;

    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    if (this.queryString) {
      try {
        const requestOptions: any = this.getRequestOptions();

        let cache: CacheStore<CacheResponse>;
        const key = JSON.stringify({
          endpoint: `${this.version}${this.SEARCH_ENDPOINT}`,
          requestOptions: requestOptions
        });
        let response = null;

        if (this.shouldRetrieveCache()) {
          cache = CacheService.getCache<CacheResponse>(schemas.search, schemas.search.stores.responses);
          const result: CacheResponse = getIsResponseCacheEnabled() ? await cache.getValue(key) : null;
          if (result && getResponseInvalidationTime(this.cacheInvalidationPeriod) > Date.now() - result.timeCached) {
            response = JSON.parse(result.response);
          }
        }

        if (!response) {
          const graph = provider.graph.forComponent(this);
          let request = graph.api(this.SEARCH_ENDPOINT).version(this.version);

          if (this.scopes && this.scopes.length) {
            request = request.middlewareOptions(prepScopes(...this.scopes));
          }

          response = await request.post(requestOptions);

          if (!equals(this.response, response)) {
            this.response = response;
          }

          if (this.shouldUpdateCache() && response) {
            cache = CacheService.getCache<CacheResponse>(schemas.search, schemas.search.stores.responses);
            cache.putValue(key, { response: JSON.stringify(response) });
          }
        }

        if (!equals(this.response, response)) {
          this.response = response;
        }
      } catch (e) {
        this.error = e;
      }

      if (this.response) {
        this.error = null;
      }
    } else {
      this.response = null;
    }
    this.isRefreshing = false;
    this.fireCustomEvent('dataChange', { response: this.response, error: this.error });
  }

  private shouldRetrieveCache(): boolean {
    return getIsResponseCacheEnabled() && this.cacheEnabled && !this.isRefreshing;
  }

  private getRequestOptions(): any {
    return {
      requests: [
        {
          entityTypes: this.entityTypes,
          query: {
            queryString: this.queryString
          },
          from: this.from ? this.from : undefined,
          size: this.size ? this.size : undefined,
          fields: this.fields ? this.fields : undefined,
          enableTopResults: this.enableTopResults ? this.enableTopResults : undefined
        }
      ]
    };
  }

  private shouldUpdateCache(): boolean {
    return getIsResponseCacheEnabled() && this.cacheEnabled;
  }
}
