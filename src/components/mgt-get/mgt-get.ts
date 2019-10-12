import { css, customElement, html, property } from 'lit-element';
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
    }
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
    type: String
  })
  public version: string = 'v1.0';

  /**
   * Maximum number of pages to get for the resource
   * default = 3
   *
   * @type {boolean}
   * @memberof MgtGet
   */
  @property({
    attribute: 'max-pages',
    type: Number
  })
  public maxPages: number = 3;

  @property({ attribute: false }) private response: any;
  @property({ attribute: false }) private error: any;
  @property({ attribute: false }) private loading: boolean = false;

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

    // TODO: handle when an attribute changes.
    //
    // Ex: load data when the name attribute changes
    // if (name === 'person-id' && oldval !== newval){
    //  this.loadData();
    // }
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
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    if (this.loading) {
      return this.renderTemplate('loading', null);
    } else if (this.error) {
      return this.renderTemplate('error', this.error);
    } else if (this.response) {
      return this.renderTemplate('default', this.response);
    }
  }

  private async loadData() {
    const provider = Providers.globalProvider;

    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    if (this.resource) {
      this.loading = true;

      let response = null;
      try {
        let request = provider.graph.client.api(this.resource).version(this.version);

        if (this.scopes && this.scopes.length) {
          request = request.middlewareOptions(prepScopes(...this.scopes));
        }

        response = await request.get();
        if (response && response.value && response.value.length && response['@odata.nextLink'] && this.maxPages) {
          let pageCount = 1;
          let page = response;

          while (pageCount < this.maxPages && page && page['@odata.nextLink']) {
            pageCount++;
            const nextResource = page['@odata.nextLink'].split(this.version)[1];
            page = await provider.graph.client
              .api(nextResource)
              .version(this.version)
              .get();
            if (page && page.value && page.value.length) {
              page.value = response.value.concat(page.value);
              response = page;
            }
          }
        }
      } catch (e) {
        this.error = e;
      }

      if (response) {
        this.response = response;
        this.error = null;
      }
    }

    this.loading = false;
  }
}
