import { ThemeCSSVariables } from '../Constants';
import { IDataSourceData } from '../models/IDataSourceData';
import { IThemeDefinition } from '../models/IThemeDefinition';
import { ITemplateService } from './ITemplateService';

/**
 * Custom attributes that can be used to display result types
 */
const ResultTypesAttributes = {
  FallbackImageUrl: 'fallback-img-url'
};

export class TemplateService implements ITemplateService {
  private _adaptiveCardsNS;
  private _markdownIt;
  private _adaptiveCardsTemplating;
  private _serializationContext;

  public async loadAdaptiveCardsResources(): Promise<void> {
    if (!this._adaptiveCardsNS) {
      // Load dynamic resources
      /*this._adaptiveCardsNS = await import(
        'adaptivecards'
      );*/

      // Initialize the serialization context for the Adaptive Cards, if needed
      if (!this._serializationContext) {
        /* const { CardObjectRegistry, GlobalRegistry, SerializationContext } = await import(
          'adaptivecards'
        );*/

        //this._serializationContext = new SerializationContext();

        const CardElementType = this._adaptiveCardsNS.CardElement;
        const ActionElementType = this._adaptiveCardsNS.Action;

        /*    const elementRegistry = new CardObjectRegistry<InstanceType<typeof CardElementType>>();
        const actionRegistry = new CardObjectRegistry<InstanceType<typeof ActionElementType>>();

        GlobalRegistry.populateWithDefaultElements(elementRegistry);
        GlobalRegistry.populateWithDefaultActions(actionRegistry);

        this._serializationContext.setElementRegistry(elementRegistry);
        this._serializationContext.setActionRegistry(actionRegistry);*/
      }

      this._adaptiveCardsNS.AdaptiveCard.onProcessMarkdown = (text: string, result) => {
        // We use Markdown here to render HTML and use web components
        const rawHtml = this._markdownIt.render(text).replace(/&lt;/g, '<').replace(/&gt;/g, '>');

        result.outputHtml = rawHtml;
        result.didProcess = true;
      };

      /* await import(
        'adaptive-expressions'
      );*/

      /* this._adaptiveCardsTemplating = await import(
        'adaptivecards-templating'
      );*/

      const MarkdownIt = await import('markdown-it');

      this._markdownIt = new MarkdownIt.default();
      this._markdownIt = new MarkdownIt();
    }
  }

  public async getFileContent(fileAbsoluteUrl: string): Promise<string> {
    const response: Response = await fetch(fileAbsoluteUrl, {
      method: 'GET',
      mode: 'cors'
    });

    if (response.ok) {
      switch (response.headers.get('Content-Type')) {
        case 'text/html' || 'text/plain':
          return await response.text();

        case 'application/json':
          return JSON.stringify(await response.json());

        default:
          return await response.text();
      }
    } else {
      throw response.statusText;
    }
  }

  public processAdaptiveCardTemplate(
    templateContent: string,
    templateContext: object,
    theme?: IThemeDefinition
  ): HTMLElement {
    const template = new this._adaptiveCardsTemplating.Template(JSON.parse(templateContent));
    const hostConfiguration = new this._adaptiveCardsNS.HostConfig(this._getHostConfiguration(theme));

    // The root context will be available in the the card implicitly
    const context = {
      $root: templateContext
    };

    const card = template.expand(context);
    const adaptiveCard = new this._adaptiveCardsNS.AdaptiveCard();
    adaptiveCard.hostConfig = hostConfiguration;

    adaptiveCard.parse(card, this._serializationContext);
    return adaptiveCard.render();
  }

  public processResultTypesFromHtml(data: IDataSourceData, templateContent: HTMLElement): HTMLElement {
    if (data?.resultTemplates) {
      // Build dictionary of available result template
      const templateDictionary = new Map(Object.entries(data.resultTemplates));

      for (const item of data.items) {
        const templateId = item.resultTemplateId;

        if (templateId) {
          const templatePayload = templateDictionary.get(templateId).body;

          // Check if item should use a result template
          if (templatePayload && templateId !== 'connectordefault') {
            // Partial match as we can"t use the complete ID due to special characters "/" and "==""
            const defaultItem: HTMLElement = templateContent.querySelector(`[id^="${item.hitId.substring(0, 15)}"]`);

            // Replace the HTML element corresponding to the item by its result type
            if (defaultItem) {
              const adaptiveCardComponent = document.createElement('mgt-adaptive-card') as any; //MgtAdaptiveCardComponent;
              adaptiveCardComponent.cardContent = JSON.stringify(templatePayload);
              adaptiveCardComponent.cardContext = item;

              // Get other settings from placeholder attribute specified in the template
              adaptiveCardComponent.fallbackImageUrl = defaultItem.getAttribute(ResultTypesAttributes.FallbackImageUrl);
              defaultItem.replaceWith(adaptiveCardComponent);
            }
          }
        }
      }
    }

    return templateContent;
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private _getHostConfiguration(theme?: IThemeDefinition): { [key: string]: any } {
    return {
      separator: {
        lineThickness: 1,
        lineColor: '#1e252b'
      },
      supportsInteractivity: true,
      fontTypes: {
        default: {
          fontFamily: theme ? theme[ThemeCSSVariables.fontFamilyPrimary] : '',
          fontSizes: {
            small: 12,
            medium: 14,
            large: 24
          },
          fontWeights: {
            bolder: 700
          }
        },
        monospace: {
          fontFamily: theme ? theme[ThemeCSSVariables.fontFamilySecondary] : '',
          fontSizes: {
            small: 12,
            medium: 14,
            large: 24
          },
          fontWeights: {
            bolder: 700
          }
        }
      },
      containerStyles: {
        default: {
          backgroundColor: '#FFFFFF',
          foregroundColors: {
            default: {
              default: '#000000'
            },
            accent: {
              default: theme ? theme[ThemeCSSVariables.colorPrimary] : ''
            },
            light: {
              default: theme ? theme[ThemeCSSVariables.colorLight] : ''
            }
          }
        },
        emphasis: {
          backgroundColor: '#DCE2E5'
        }
      },
      adaptiveCard: {
        allowCustomStyle: true
      },
      spacing: {
        default: 6,
        medium: 8,
        large: 16
      }
    };
  }
}
