/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { DriveItem } from '@microsoft/microsoft-graph-types';
import { customElement, html, property, TemplateResult } from 'lit-element';
import { styles } from './mgt-file-css';
import { MgtTemplatedComponent, Providers, ProviderState } from '@microsoft/mgt-element';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';
import { getDriveItem } from '../../graph/graph.files';
import { getRelativeDisplayDate } from '../../utils/Utils';
import { ViewType } from '../../graph/types';
import { deleteTodoTaskList } from '../mgt-todo/graph.todo';

/**
 * The File component is used to represent an individual file/folder from OneDrive or SharePoint by displaying information such as the file/folder name, an icon indicating the file type, and other properties such as the author, last modified date, or other details selected by the developer.
 *
 * @export
 * @class MgtFile
 * @extends {MgtTemplatedComponent}
 */
@customElement('mgt-file')
export class MgtFile extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * allows developer to provide query for a file
   *
   * @type {string}
   * @memberof MgtFile
   */
  @property({
    attribute: 'file-query'
  })
  public get fileQuery(): string {
    return this._fileQuery;
  }
  public set fileQuery(value: string) {
    if (value === this._fileQuery) {
      return;
    }

    this._fileQuery = value;
    this.requestStateUpdate();
    // this.requestUpdate('fileQuery');
  }

  /**
   * object containing Graph details on driveItem
   *
   * @type {string}
   * @memberof MgtFile
   */
  @property({
    attribute: 'drive-item',
    type: Object
  })
  public driveItem: DriveItem;

  /**
   * Sets the property of the file to use for the first line of text.
   * Default is file name
   *
   * @type {string}
   * @memberof MgtFile
   */
  @property({ attribute: 'line1-property' }) public line1Property: string;

  /**
   * Sets the property of the file to use for the second line of text.
   * Default is last modified date time
   *
   * @type {string}
   * @memberof MgtFile
   */
  @property({ attribute: 'line2-property' }) public line2Property: string;

  /**
   * Sets the property of the file to use for the second line of text.
   * Default is file size
   *
   * @type {string}
   * @memberof MgtFile
   */
  @property({ attribute: 'line3-property' }) public line3Property: string;

  /**
   * Sets what data to be rendered (file icon only, oneLine, twoLines threeLines).
   * Default is 'threeLines'.
   *
   * @type {ViewType}
   * @memberof MgtFile
   */
  @property({
    converter: value => {
      if (!value || value.length === 0) {
        return ViewType.threelines;
      }

      value = value.toLowerCase();

      if (typeof ViewType[value] === 'undefined') {
        return ViewType.threelines;
      } else {
        return ViewType[value];
      }
    }
  })
  public view: ViewType;

  private _fileQuery: string;

  constructor() {
    super();
    this.line1Property = 'name';
    this.line2Property = 'lastModifiedDateTime';
    this.line3Property = 'size';
    this.view = ViewType.threelines;
  }

  public render() {
    if (!this.driveItem && this.isLoadingState) {
      return this.renderLoading();
    }

    if (!this.driveItem) {
      return this.renderNoData();
    }

    const file = this.driveItem;
    let fileTemplate = this.renderTemplate('default', { file });

    if (!fileTemplate) {
      const fileDetailsTemplate: TemplateResult = this.renderDetails(file);
      const fileTypeIconTemplate: TemplateResult = this.renderFileTypeIcon();

      fileTemplate = html`
        <div class="item">
          ${fileTypeIconTemplate} ${fileDetailsTemplate}
        </div>
      `;
    }

    return html`
      ${fileTemplate}
    `;
  }

  /**
   * Render the loading state
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtFile
   */
  protected renderLoading(): TemplateResult {
    return this.renderTemplate('loading', null) || html``;
  }

  /**
   * Render the state when no data is available
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtFile
   */
  protected renderNoData(): TemplateResult {
    return this.renderTemplate('no-data', null) || html``;
  }

  /**
   * Render the file type icon
   *
   * @protected
   * @param {string} [iconSrc]
   * @memberof MgtFile
   */
  protected renderFileTypeIcon(): TemplateResult {
    return html`
      <div class="item__file-type-icon">
        ${getSvg(SvgIcon.File)}
      </div>
    `;
  }

  /**
   * Render the file type icon
   *
   * @protected
   * @param {MicrosoftGraph.DriveItem} [driveItem]
   * @memberof MgtFile
   */
  protected renderDetails(driveItem: DriveItem): TemplateResult {
    if (!driveItem || this.view === ViewType.image) {
      return html``;
    }

    const details: TemplateResult[] = [];

    if (this.view > ViewType.image) {
      const text = this.getTextFromProperty(driveItem, this.line1Property);
      if (text) {
        details.push(html`
          <div class="line1" aria-label="${text}">${text}</div>
        `);
      }
    }

    if (this.view > ViewType.oneline) {
      const text = this.getTextFromProperty(driveItem, this.line2Property);
      if (text) {
        details.push(html`
          <div class="line2" aria-label="${text}">${text}</div>
        `);
      }
    }

    if (this.view > ViewType.twolines) {
      const text = this.getTextFromProperty(driveItem, this.line3Property);
      if (text) {
        details.push(html`
          <div class="line3" aria-label="${text}">${text}</div>
        `);
      }
    }

    return html`
      <div class="item__details">
        ${details}
      </div>
    `;
  }

  protected async loadState() {
    const provider = Providers.globalProvider;
    if (!provider || provider.state === ProviderState.Loading) {
      return;
    }

    if (provider.state === ProviderState.SignedOut) {
      this.driveItem = null;
      return;
    }

    const graph = provider.graph.forComponent(this);
    const driveItem = await getDriveItem(graph, this.fileQuery);
    this.driveItem = driveItem;
  }

  private getFileTypeIcon() {
    return;
    // graph call to get file type
    // determine which icon to render based on file type
  }

  private getTextFromProperty(driveItem: DriveItem, properties: string) {
    if (!properties || properties.length === 0) {
      return null;
    }

    const propertyList = properties.trim().split(',');
    let text;
    let i = 0;

    // convert date time
    const lastModifiedDateTime = new Date(driveItem.lastModifiedDateTime);
    const relativeDateString = getRelativeDisplayDate(lastModifiedDateTime);

    // convert size to mb
    const sizeInMb = (driveItem.size / (1024 * 1024)).toFixed(2);

    while (!text && i < propertyList.length) {
      const current = propertyList[i].trim();
      switch (current) {
        case 'size':
          text = `Size: ${sizeInMb}MB`;
          break;
        case 'lastModifiedDateTime':
          text = `Modified ${relativeDateString}`;
          break;
        default:
          text = driveItem[current];
      }
      i++;
    }

    return text;
  }
}
