import { FileFormat, MgtConnectableComponent, TemplateService } from '@microsoft/mgt-element';
import { css, html, PropertyValues } from 'lit';
import { unsafeHTML } from 'lit/directives/unsafe-html.js';
import { styles as tailwindStyles } from '../../../../styles/tailwind-styles-css';
import { customElement, property, state } from 'lit/decorators.js';

/**
 * Process adaptive card content from an external file
 */
@customElement('mgt-adaptive-card')
export class MgtAdaptiveCardComponent extends MgtConnectableComponent {
  /**
   * The file URL to fetch
   */
  @property({ type: String, attribute: 'url' })
  fileUrl: string;

  /**
   * The file format to load
   */
  @property({ type: String, attribute: 'format' })
  fileFormat: FileFormat = FileFormat.Json;

  /**
   * The fallback image URL
   */
  @property({ type: String, attribute: 'fallback-img-url' })
  fallbackImageUrl: string;

  /**
   * The data context to use to render the card
   */
  @property({
    type: Object,
    attribute: 'context',
    converter: {
      fromAttribute: value => {
        try {
          return JSON.parse(value);
        } catch {
          return null;
        }
      }
    }
  })
  cardContext: object;

  /**
   * The raw adaptive card content as string (i.e. JSON stringified)
   */
  @property({ type: String, attribute: 'content' })
  cardContent: string;

  /**
   * The file content to display
   */
  @state()
  content: string;

  constructor() {
    super();
    this.onImageError = this.onImageError.bind(this);
  }

  render() {
    if (this.content) {
      return html`<div class="animate-fadein">${unsafeHTML(this.content)}</div>`;
    } else {
      return html`
                <div class="flex justify-center space-x-2 opacity-75" >
                    <div style="animation-delay: 0.1s" class="bg-primary p-1 w-1 h-1 rounded-full animate-bounce"></div>
                    <div style="animation-delay: 0.2s" class="bg-primary p-1 w-1 h-1 rounded-full animate-bounce"></div>
                    <div style="animation-delay: 0.3s" class="bg-primary p-1 w-1 h-1 rounded-full animate-bounce"></div>
                </div>
            `;
    }
  }

  static get styles() {
    return [
      css`
            :host  {               

                > .ac-container{
                    padding: 8px !important;
                }

                .ac-anchor {
                    color: var(--text-color);
                    text-decoration: none;
                    transition-property: all;
                    transition-timing-function: cubic-bezier(0.4, 0, 0.2, 1);
                    transition-duration: 150ms;
                }
                .ac-textBlock{
                    white-space: normal !important;
                    overflow-wrap: break-word !important;
                }

                .ac-anchor:hover{
                    color: var(--ubi365-colorPrimary);
                }

                .ac-anchor:focus, .ac-anchor:focus-visible{
                    color: var(--ubi365-colorPrimary);
                }

                .ac-container#ubiSearchTitle > .ac-textBlock{
                    overflow: unset !important;
                }

                .ac-container#ubiSearchTitle > .ac-textBlock > p{
                    overflow: unset !important;
                }

                .ac-container#ubiSearchSite > .ac-textBlock{
                    overflow: unset !important;
                }
                .ac-container#ubiSearchSite > .ac-textBlock > p{
                    overflow: unset !important;
                }

                .ac-container#ubiSearchSite > .ac-textBlock > p > a{
                    display: inline-block;
                    padding-left: 16px!important;
                    padding-right: 16px!important;
                    padding-top: 4px!important;
                    padding-bottom: 4px!important;
                    border-radius: 25px;
                    background-color: var(--topic-background-color) !important;
                    text-decoration: none;
                    color: var(--text-color);
                    transition-property: all;
                    transition-timing-function: cubic-bezier(0.4, 0, 0.2, 1);
                    transition-duration: 150ms;
                }

                .ac-container#ubiSearchSite > .ac-textBlock > p > a:hover{
                    background-color: var(--topic-hover-color) !important;
                }
                .ac-container#ubiSearchSite > .ac-textBlock > p > a:focus, .ac-container#ubiSearchSite > .ac-textBlock > p > a:focus-visible{
                    background-color: var(--topic-focus-color) !important;
                }
                #ubiSearchSummary{
                    font-size: 1rem !important;
                    line-height: 1.5rem !important;
                }
            }                 
        `,
      tailwindStyles
    ];
  }

  public disconnectedCallback(): void {
    super.disconnectedCallback();
  }

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  protected updated(changedProperties: PropertyValues<this>): void {
    // Set default fallback image if not found
    if (this.fallbackImageUrl) {
      this.renderRoot.querySelectorAll('img').forEach(el => {
        if (el && !this._eventHanlders.get(JSON.stringify(el))) {
          this._eventHanlders.set(JSON.stringify(el), { component: el, event: 'error', handler: this.onImageError });
          el.addEventListener('error', this.onImageError);
        }
      });
    }
  }

  private onImageError(event: Event) {
    (event.target as HTMLImageElement).src = this.fallbackImageUrl;

    // To avoid endless loop
    event.target.removeEventListener('error', this.onImageError);
    this._eventHanlders.delete(JSON.stringify(event.target));
    (event.target as HTMLImageElement).onerror = null;
  }

  public async connectedCallback(): Promise<void> {
    const io = new IntersectionObserver(async data => {
      if (data[0].isIntersecting) {
        await this._processAdaptiveCard();
        io.disconnect();
      }
    });

    io.observe(this);

    super.connectedCallback();
  }

  private async _processAdaptiveCard() {
    const templateService = new TemplateService();
    await templateService.loadAdaptiveCardsResources();

    let contentToProcess: string;

    if (this.cardContent) {
      contentToProcess = this.cardContent;
    } else if (this.fileUrl) {
      // Get the file raw content according to the type
      const fileContent = await templateService.getFileContent(this.fileUrl);

      if (this.fileFormat === FileFormat.Json) {
        contentToProcess = fileContent;
      }
    }

    if (this.cardContext) {
      const htmlContent: HTMLElement = templateService.processAdaptiveCardTemplate(
        contentToProcess,
        this.cardContext,
        null
      );
      this.content = htmlContent.innerHTML;
    } else {
      this.content = this.cardContent;
    }
  }
}
