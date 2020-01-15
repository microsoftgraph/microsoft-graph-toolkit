/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, property } from 'lit-element';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import { prepScopes } from '../../utils/GraphHelpers';
import { MgtTemplatedComponent } from '../templatedComponent';

/**
 * Custom element for making Microsoft Graph get queries
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
   * @type {*}
   * @memberof MgtGet
   */
  @property({ attribute: false }) public response: any;

  /**
   *
   * Gets or sets the error (if any) of the request
   * @type {*}
   * @memberof MgtGet
   */
  @property({ attribute: false }) public error: any;

  @property({ attribute: false }) private loading: boolean = false;

  private hasFirstUpdated: boolean = false;
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

    if (this.hasFirstUpdated) {
      this.response = null;
      this.loadData();
    }
  }

  /**
   * Invoked when the element is first updated. Implement to perform one time
   * work on the element after update.
   *
   * Setting properties inside this method will trigger the element to update
   * again after this update cycle completes.
   *
   * * @param _changedProperties Map of changed properties with old values
   */
  public firstUpdated() {
    Providers.onProviderUpdated(() => this.loadData());
    this.loadData();
    this.hasFirstUpdated = true;
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    if (this.error) {
      return this.renderTemplate('error', this.error);
      // tslint:disable-next-line: no-string-literal
    } else if (this.templates['value'] && this.response && this.response.value) {
      if (Array.isArray(this.response.value)) {
        let loading = null;
        if (this.loading && !this.isPolling) {
          loading = this.renderTemplate('loading', null);
        }
        return html`
          ${this.response.value.map(v => this.renderTemplate('value', v, v.id))} ${loading}
        `;
      } else {
        return this.renderTemplate('value', this.response);
      }
    } else if (this.response) {
      return this.renderTemplate('default', this.response);
    } else if (this.loading) {
      return this.renderTemplate('loading', null);
    }
  }

  private async loadData() {
    const provider = Providers.globalProvider;

    this.error = null;

    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    if (this.resource) {
      this.loading = true;

      try {
        const graph = provider.graph.forComponent(this);

        // if we had a response earlier with a delta link, use it instead
        const uri =
          this.response && this.response['@odata.deltaLink'] ? this.response['@odata.deltaLink'] : this.resource;
        let request = graph.client.api(uri).version(this.version);

        if (this.scopes && this.scopes.length) {
          request = request.middlewareOptions(prepScopes(...this.scopes));
        }

        const response = await request.get();

        if (this.response && Array.isArray(this.response.value) && Array.isArray(response.value)) {
          response.value = this.response.value.concat(response.value);
        }

        this.response = response;

        // get more pages if there are available
        if (this.response && Array.isArray(this.response.value) && this.response['@odata.nextLink']) {
          let pageCount = 1;
          let page = this.response;

          while (
            (pageCount < this.maxPages || this.maxPages <= 0 || this.pollingRate) &&
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
              page.value = this.response.value.concat(page.value);
              this.response = page;
            }
          }
        }
      } catch (e) {
        this.error = e;
      }

      if (this.response) {
        this.error = null;

        if (this.pollingRate) {
          setTimeout(async () => {
            this.isPolling = true;
            await this.loadData();
            this.isPolling = false;
          }, this.pollingRate);
        }
      }
      this.loading = false;
    } else {
      this.response = null;
    }

    this.fireCustomEvent('dataChange', { response: this.response, error: this.error });
  }
}
