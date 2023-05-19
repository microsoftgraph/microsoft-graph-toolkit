import { IDataSourceData } from '../models/IDataSourceData';
import { IThemeDefinition } from '../models/IThemeDefinition';

export enum FileFormat {
  Text = 'text',
  Json = 'json'
}

export interface ITemplateService {
  /**
   * Update the HTML element with corresponding result types for items (i.e. node with id attribute equals to "hitId")
   * @param data the data source data containing the items
   * @param templateContent the template content as HTML
   * @param theme the theme to apply
   * @returns the updated HTML element with result types
   */
  processResultTypesFromHtml(
    data: IDataSourceData,
    templateContent: HTMLElement,
    theme?: IThemeDefinition
  ): HTMLElement;

  /**
   * Process the adaptive card with data from the context
   * @param templateContent the card content as stringified JSON
   * @param templateContext the card context for data binding
   * @param theme the theme to apply
   * @returns the processed HTML as raw string
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  processAdaptiveCardTemplate(
    templateContent: string,
    templateContext: { [key: string]: any },
    theme?: IThemeDefinition
  ): HTMLElement;

  /**
   * Load adaptive cards resources on the page dynamically
   */
  loadAdaptiveCardsResources(): Promise<void>;

  /**
   * Gets the external file content from the specified URL using Microsoft Graph
   * @param fileUrl
   * @returns the file raw content as string
   */
  getFileContent(fileUrl: string): Promise<string>;
}
