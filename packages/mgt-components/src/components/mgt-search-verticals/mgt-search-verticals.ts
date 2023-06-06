import { EventConstants, customElement } from '@microsoft/mgt-element';
import { property } from 'lit/decorators.js';
import { html } from 'lit';
import { IDataVerticalConfiguration } from '@microsoft/mgt-element/src/models/IDataVerticalConfiguration';
import { ISearchVerticalEventData, MgtConnectableComponent, PageOpenBehavior } from '@microsoft/mgt-element';
import { repeat } from 'lit/directives/repeat.js';
import { isEmpty, isEqual } from 'lodash-es';
import { styles } from './mgt-search-verticals-css';
import { styles as tailwindStyles } from '../../styles/tailwind-styles-css';
import { fluentTab, fluentTabPanel, fluentTabs, provideFluentDesignSystem } from '@fluentui/web-components';
import { UrlHelper } from '@microsoft/mgt-element';

@customElement('search-verticals')
export class MgtSearchVerticalsComponent extends MgtConnectableComponent {
  /**
   * The configured search verticals
   */
  @property({
    type: Object,
    attribute: 'settings',
    converter: {
      fromAttribute: value => {
        return JSON.parse(value) as IDataVerticalConfiguration[];
      }
    }
  })
  public verticals: IDataVerticalConfiguration[];

  /**
   * The query string parameter name to use to select a vertical tab by default
   */
  @property({ type: String, attribute: 'default-query-string-param' })
  public defaultVerticalQueryStringParam = 'v';

  @property({ type: String, attribute: 'selected-key', reflect: true })
  public selectedVerticalKey: string;

  constructor() {
    super();
    this.onVerticalSelected = this.onVerticalSelected.bind(this);

    // Register fluent tabs (as scoped elements)
    provideFluentDesignSystem().register(fluentTab(), fluentTabs(), fluentTabPanel());
  }

  public render() {
    return html`
        <div class="px-2.5">
          <div class="max-w-7xl ml-auto mr-auto mb-8">
            <fluent-tabs activeid=${this.selectedVerticalKey}>
              ${repeat(
                this.verticals,
                vertical => vertical.key,
                vertical => {
                  return html`
                    <fluent-tab 
                      id=${vertical.key}
                      data-name=${this.getLocalizedString(vertical.tabName)}
                      data-key=${vertical.key}
                      @click=${() => {
                        this.onVerticalSelected(vertical);
                      }}
                    >
                      ${this.getLocalizedString(vertical.tabName)}
                    </fluent-tab>
                    <fluent-tab-panel id=${`${vertical.key}Panel`}>
                      </fluent-tab-panel>
                  `;
                }
              )}
            </fluent-tabs>
          </div>
        </div>
      `;
  }

  static get styles() {
    return [
      styles, // Allow component to use them CSS variables from design. The component is a first level component so it is OK to define theme variables here
      tailwindStyles // Use Tailwind CSS classes
    ];
  }

  public connectedCallback(): void {
    this.handleQueryStringChange();
    this.initializeDefaultValue();

    return super.connectedCallback();
  }

  /**
   * Select the vertical and notifies all subscribers
   * @param verticalKey the vertical key to select
   */
  public selectVertical(verticalKey: string) {
    if (!isEmpty(verticalKey)) {
      this.selectedVerticalKey = verticalKey;

      // Update the query string parameter if already present in the URL
      const verticalKeyParam = UrlHelper.getQueryStringParam(
        this.defaultVerticalQueryStringParam,
        window.location.href
      );
      if (
        this.defaultVerticalQueryStringParam &&
        verticalKeyParam &&
        !isEqual(this.selectedVerticalKey, verticalKeyParam)
      ) {
        window.history.pushState(
          {},
          '',
          UrlHelper.addOrReplaceQueryStringParam(
            window.location.href,
            this.defaultVerticalQueryStringParam,
            this.selectedVerticalKey
          )
        );
      }

      this.fireCustomEvent(EventConstants.SEARCH_VERTICAL_EVENT, {
        selectedVertical: this.verticals.filter(v => v.key === this.selectedVerticalKey)[0]
      } as ISearchVerticalEventData);
    }
  }

  /**
   * Handler when a vertical is slected by an user
   * @param vertical the current selected vertical
   */
  private onVerticalSelected(vertical: IDataVerticalConfiguration): void {
    if (vertical.isLink && vertical.linkUrl) {
      window.open(vertical.linkUrl, PageOpenBehavior.NewTab ? '_blank' : '');
    } else {
      this.selectVertical(vertical.key);
    }
  }

  /**
   * Initialize the default vertical value according to settings
   */
  private initializeDefaultValue() {
    if (!this.defaultVerticalQueryStringParam) {
      this.selectedVerticalKey = this.selectedVerticalKey ? this.selectedVerticalKey : this.verticals[0].key;
    } else {
      // Get vertical corresponding to the query string URL parameter
      const defaultQueryVal = UrlHelper.getQueryStringParam(
        this.defaultVerticalQueryStringParam,
        window.location.href.toLowerCase()
      );

      if (defaultQueryVal) {
        const defaultSelected: IDataVerticalConfiguration[] = this.verticals.filter(
          v => v.key.toLowerCase() === decodeURIComponent(defaultQueryVal.toLowerCase())
        );
        if (defaultSelected.length === 1) {
          this.selectedVerticalKey = defaultSelected[0].key;
        }
      } else {
        this.selectedVerticalKey = this.selectedVerticalKey ? this.selectedVerticalKey : this.verticals[0].key;
      }
    }
  }

  /**
   * Subscribes to URL query string change events using windows state
   */
  private handleQueryStringChange() {
    if (this.defaultVerticalQueryStringParam) {
      // Will fire on browser back/forward
      window.onpopstate = () => {
        const verticalKeyParam = UrlHelper.getQueryStringParam(
          this.defaultVerticalQueryStringParam,
          window.location.href
        );
        if (this.selectedVerticalKey !== verticalKeyParam) this.selectVertical(verticalKeyParam);
      };
    }
  }
}
