/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, property } from 'lit-element';
import { Providers, ProviderState } from '@microsoft/mgt-element';
import { prepScopes } from '../../utils/GraphHelpers';
import { equals } from '../../utils/Utils';
import { MgtTemplatedComponent } from '../templatedComponent';

/**
 * Custom element for making Microsoft Graph get queries
 *
 * @fires dataChange - Fired when data changes
 *
 * @export
 * @class mgt-get
 * @extends {MgtTemplatedComponent}
 */
@customElement('mgt-get')
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
   * The resource to get
   *
   * @type {string[]}
   * @memberof MgtGet
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
   * @memberof MgtGet
   */
  @property({
    attribute: 'version',
    reflect: true,
    type: String
  })
  public version: string = 'v1.0';

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
  public maxPages: number = 3;

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
  public pollingRate: number = 0;

  /**
   * Gets or sets the response of the request
   *
   * @type any
   * @memberof MgtGet
   */
  @property({ attribute: false }) public response: any;

  /**
   *
   * Gets or sets the error (if any) of the request
   * @type any
   * @memberof MgtGet
   */
  @property({ attribute: false }) public error: any;

  private isPolling: boolean = false;

  /**
   * Synchronizes property values when attributes change.
   *
   * @param {*} name
   * @param {*} oldValue
   * @param {*} newValue
   * @memberof MgtPersonCard
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
   * @memberof MgtGet
   */
  public refresh(hardRefresh = false) {
    this.requestStateUpdate(hardRefresh);
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    if (this.isLoadingState && !this.isPolling) {
      return this.renderTemplate('loading', null);
    } else if (this.error) {
      return this.renderTemplate('error', this.error);
      // tslint:disable-next-line: no-string-literal
    } else if (this.hasTemplate('value') && this.response && this.response.value) {
      let valueContent;

      if (Array.isArray(this.response.value)) {
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

      // tslint:disable-next-line: no-string-literal
      if (this.hasTemplate('default')) {
        const defaultContent = this.renderTemplate('default', this.response);

        // tslint:disable-next-line: no-string-literal
        if ((this.templates['value'] as any).templateOrder > (this.templates['default'] as any).templateOrder) {
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
        let uri = this.resource;
        let isDeltaLink = false;

        // if we had a response earlier with a delta link, use it instead
        if (this.response && this.response['@odata.deltaLink']) {
          uri = this.response['@odata.deltaLink'];
          isDeltaLink = true;
        } else {
          isDeltaLink = new URL(uri, 'https://graph.microsoft.com').pathname.endsWith('delta');
        }

        const graph = provider.graph.forComponent(this);
        let request = graph.api(uri).version(this.version);

        if (this.scopes && this.scopes.length) {
          request = request.middlewareOptions(prepScopes(...this.scopes));
        }

        let response = await request.get();

        if (isDeltaLink && this.response && Array.isArray(this.response.value) && Array.isArray(response.value)) {
          response.value = this.response.value.concat(response.value);
        }

        if (!this.isPolling && !equals(this.response, response)) {
          this.response = response;
        }

        // get more pages if there are available
        if (response && Array.isArray(response.value) && response['@odata.nextLink']) {
          let pageCount = 1;
          let page = response;

          while (
            (pageCount < this.maxPages || this.maxPages <= 0 || (isDeltaLink && this.pollingRate)) &&
            page &&
            page['@odata.nextLink']
          ) {
            pageCount++;
            const nextResource = page['@odata.nextLink'].split(this.version)[1];
            page = await graph.client
              .api(nextResource)
              .version(this.version)
              .get();
            if (page && page.value && page.value.length) {
              page.value = response.value.concat(page.value);
              response = page;
              if (!this.isPolling) {
                this.response = response;
              }
            }
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

        if (this.pollingRate) {
          setTimeout(async () => {
            this.isPolling = true;
            await this.loadState();
            this.isPolling = false;
          }, this.pollingRate);
        }
      }
    } else {
      this.response = null;
    }

    this.fireCustomEvent('dataChange', { response: this.response, error: this.error });
  }
}
