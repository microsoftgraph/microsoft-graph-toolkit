/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html, TemplateResult } from 'lit';
import { property } from 'lit/decorators.js';
import {
  CacheService,
  CacheStore,
  equals,
  MgtTemplatedComponent,
  prepScopes,
  Providers,
  ProviderState,
  customElement,
  CollectionResponse
} from '@microsoft/mgt-element';

import { getPhotoForResource } from '../../graph/graph.photos';
import { getDocumentThumbnail } from '../../graph/graph.files';
import { schemas } from '../../graph/cacheStores';
import { CacheResponse } from '../CacheResponse';
import { Entity } from '@microsoft/microsoft-graph-types';

/**
 * Simple holder type for an image
 */
interface ImageValue {
  image: string;
}

/**
 * A type guard to check if a value is a collection response
 *
 * @param value {*} the value to check
 * @returns {boolean} true if the value is a collection response
 */
export const isCollectionResponse = (value: unknown): value is CollectionResponse<unknown> =>
  Array.isArray((value as CollectionResponse<unknown>)?.value);

/**
 * Enumeration to define what types of query are available
 *
 * @export
 * @enum {string}
 */
export enum ResponseType {
  /**
   * Fetches a call as JSON
   */
  json = 'json',

  /**
   * Fetches a call as image
   */
  image = 'image'
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
 * Holder type emitted with the dataChange event
 */
export interface DataChangedDetail {
  response?: CollectionResponse<Entity>;
  error?: object;
}

/**
 * Custom element for making Microsoft Graph get queries
 *
 * @fires {CustomEvent<DataChangedDetail>} dataChange - Fired when data changes
 *
 * @export
 * @class mgt-get
 * @extends {MgtTemplatedComponent}
 */
@customElement('get')
export class MgtGet extends MgtTemplatedComponent {
  /**
   * The resource to get
   *
   * @type {string}
   * @memberof MgtGet
   */
  @property({
    attribute: 'resource',
    reflect: true,
    type: String
  })
  public resource: string;

  /**
   * The scopes to request
   *
   * @type {string[]}
   * @memberof MgtGet
   */
  @property({
    attribute: 'scopes',
    converter: (value, _type) => {
      return value ? value.toLowerCase().split(',') : null;
    },
    reflect: true
  })
  public scopes: string[] = [];

  /**
   * Api version to use for request
   *
   * @type {string}
   * @memberof MgtGet
   */
  @property({
    attribute: 'version',
    reflect: true,
    type: String
  })
  public version = 'v1.0';

  /**
   * Type of response
   * Default = json
   * Supported values = json, image
   *
   * @type {ResponseType}
   * @memberof MgtGet
   */
  @property({
    attribute: 'type',
    reflect: true,
    type: ResponseType
  })
  public type: ResponseType = ResponseType.json;

  /**
   * Maximum number of pages to get for the resource
   * default = 3
   * if <= 0, all pages will be fetched
   *
   * @type {number}
   * @memberof MgtGet
   */
  @property({
    attribute: 'max-pages',
    reflect: true,
    type: Number
  })
  public maxPages = 3;

  /**
   * Number of milliseconds to poll the delta API and
   * update the response. Set to positive value to enable
   *
   * @type {number}
   * @memberof MgtGet
   */
  @property({
    attribute: 'polling-rate',
    reflect: true,
    type: Number
  })
  public pollingRate = 0;

  /**
   * Enables cache on the response from the specified resource
   * default = false
   *
   * @type {boolean}
   * @memberof MgtGet
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
   * @memberof MgtGet
   */
  @property({
    attribute: 'cache-invalidation-period',
    type: Number
  })
  public cacheInvalidationPeriod = 0;

  /**
   * Gets or sets the response of the request
   *
   * @type any
   * @memberof MgtGet
   */
  @property({ attribute: false }) public response: CollectionResponse<Entity> | Entity | ImageValue;

  /**
   *
   * Gets or sets the error (if any) of the request
   *
   * @type any
   * @memberof MgtGet
   */
  @property({ attribute: false }) public error: object;

  private isPolling = false;
  private isRefreshing = false;

  /**
   * Synchronizes property values when attributes change.
   *
   * @param {*} name
   * @param {*} oldValue
   * @param {*} newValue
   * @memberof MgtPersonCard
   */
  public attributeChangedCallback(name, oldval, newval) {
    // eslint-disable-next-line @typescript-eslint/no-unsafe-argument
    super.attributeChangedCallback(name, oldval, newval);
    void this.requestStateUpdate();
  }

  /**
   * Refresh the data
   *
   * @param {boolean} [hardRefresh=false]
   * if false (default), the component will only update if the data changed
   * if true, the data will be first cleared and reloaded completely
   * @memberof MgtGet
   */
  public refresh(hardRefresh = false) {
    this.isRefreshing = true;
    if (hardRefresh) {
      this.clearState();
    }
    void this.requestStateUpdate(hardRefresh);
  }

  /**
   * Clears state of the component
   *
   * @protected
   * @memberof MgtGet
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
    if (this.isLoadingState && !this.response) {
      return this.renderTemplate('loading', null);
    } else if (this.error) {
      return this.renderTemplate('error', this.error);
      // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access, @typescript-eslint/dot-notation
    } else if (this.hasTemplate('value') && isCollectionResponse(this.response)) {
      let valueContent: TemplateResult;

      if (isCollectionResponse(this.response)) {
        let loading = null;
        if (this.isLoadingState && !this.isPolling) {
          loading = this.renderTemplate('loading', null);
        }
        valueContent = html`
          ${this.response.value.map(v => this.renderTemplate('value', v, v.id))} ${loading}
        `;
      } else {
        valueContent = this.renderTemplate('value', this.response);
      }

      if (this.hasTemplate('default')) {
        const defaultContent = this.renderTemplate('default', this.response);

        // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access, @typescript-eslint/dot-notation
        if (this.templates['value'].templateOrder > this.templates['default'].templateOrder) {
          return html`
            ${defaultContent}${valueContent}
          `;
        } else {
          return html`
            ${valueContent}${defaultContent}
          `;
        }
      } else {
        return valueContent;
      }
    } else if (this.response) {
      return this.renderTemplate('default', this.response) || html``;
    } else if (this.hasTemplate('no-data')) {
      return this.renderTemplate('no-data', null);
    } else {
      return html``;
    }
  }

  /**
   * load state into the component.
   *
   * @protected
   * @returns
   * @memberof MgtGet
   */
  protected async loadState() {
    const provider = Providers.globalProvider;

    this.error = null;

    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    if (this.resource) {
      try {
        let cache: CacheStore<CacheResponse>;
        const key = `${this.version}${this.resource}`;
        let response: Entity | CollectionResponse<Entity> | ImageValue = null;

        if (this.shouldRetrieveCache()) {
          cache = CacheService.getCache<CacheResponse>(schemas.get, schemas.get.stores.responses);
          const result: CacheResponse = getIsResponseCacheEnabled() ? await cache.getValue(key) : null;
          if (result && getResponseInvalidationTime(this.cacheInvalidationPeriod) > Date.now() - result.timeCached) {
            response = JSON.parse(result.response) as CollectionResponse<Entity>;
          }
        }

        if (!response) {
          let uri = this.resource;
          let isDeltaLink = false;

          // if we had a response earlier with a delta link, use it instead
          if (this.response?.['@odata.deltaLink']) {
            uri = this.response['@odata.deltaLink'] as string;
            isDeltaLink = true;
          } else {
            // TODO: Check this against the base url for the cloud in use.
            isDeltaLink = new URL(uri, 'https://graph.microsoft.com').pathname.endsWith('delta');
          }

          const graph = provider.graph.forComponent(this);
          let request = graph.api(uri).version(this.version);

          if (this.scopes?.length) {
            request = request.middlewareOptions(prepScopes(...this.scopes));
          }

          if (this.type === ResponseType.json) {
            response = (await request.get()) as CollectionResponse<Entity> | Entity;

            if (isDeltaLink && isCollectionResponse(this.response) && isCollectionResponse(response)) {
              const responseValues: Entity[] = response.value;
              response.value = this.response.value.concat(responseValues);
            }

            if (!this.isPolling && !equals(this.response, response)) {
              this.response = response;
            }

            // get more pages if there are available
            if (isCollectionResponse(response) && response['@odata.nextLink']) {
              let pageCount = 1;
              let page = response;

              while (
                (pageCount < this.maxPages || this.maxPages <= 0 || (isDeltaLink && this.pollingRate)) &&
                page?.['@odata.nextLink']
              ) {
                pageCount++;
                const nextResource = (page['@odata.nextLink'] as string).split(this.version)[1];
                page = (await graph.client.api(nextResource).version(this.version).get()) as CollectionResponse<Entity>;
                if (page?.value?.length) {
                  page.value = response.value.concat(page.value);
                  response = page;
                  if (!this.isPolling) {
                    this.response = response;
                  }
                }
              }
            }
          } else {
            if (this.resource.indexOf('/photo/$value') === -1 && this.resource.indexOf('/thumbnails/') === -1) {
              throw new Error('Only /photo/$value and /thumbnails/ endpoints support the image type');
            }

            let image: string;
            if (this.resource.indexOf('/photo/$value') > -1) {
              // Sanitizing the resource to ensure getPhotoForResource gets the right format
              const sanitizedResource = this.resource.replace('/photo/$value', '');
              const photoResponse = await getPhotoForResource(graph, sanitizedResource, this.scopes);
              if (photoResponse) {
                image = photoResponse.photo;
              }
            } else if (this.resource.indexOf('/thumbnails/') > -1) {
              const imageResponse = await getDocumentThumbnail(graph, this.resource, this.scopes);
              if (imageResponse) {
                image = imageResponse.thumbnail;
              }
            }

            if (image) {
              response = {
                image
              };
            }
          }

          if (this.shouldUpdateCache() && response) {
            cache = CacheService.getCache<CacheResponse>(schemas.get, schemas.get.stores.responses);
            await cache.putValue(key, { response: JSON.stringify(response) });
          }
        }

        if (!equals(this.response, response)) {
          this.response = response;
        }
      } catch (e: unknown) {
        this.error = e as object;
      }

      if (this.response) {
        this.error = null;

        if (this.pollingRate) {
          setTimeout(() => {
            this.isPolling = true;
            void this.loadState().finally(() => {
              this.isPolling = false;
            });
          }, this.pollingRate);
        }
      }
    } else {
      this.response = null;
    }
    this.isRefreshing = false;
    this.fireCustomEvent('dataChange', { response: this.response, error: this.error });
  }

  private shouldRetrieveCache(): boolean {
    return getIsResponseCacheEnabled() && this.cacheEnabled && !(this.isRefreshing || this.isPolling);
  }

  private shouldUpdateCache(): boolean {
    return getIsResponseCacheEnabled() && this.cacheEnabled;
  }
}
