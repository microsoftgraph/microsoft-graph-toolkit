/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { DriveItem } from '@microsoft/microsoft-graph-types';
import { html, TemplateResult } from 'lit';
import { property } from 'lit/decorators.js';
import { styles } from './mgt-file-css';
import { MgtTemplatedComponent, Providers, ProviderState, customElement } from '@microsoft/mgt-element';
import {
  getDriveItemById,
  getDriveItemByPath,
  getDriveItemByQuery,
  getGroupDriveItemById,
  getGroupDriveItemByPath,
  getListDriveItemById,
  getMyDriveItemById,
  getMyDriveItemByPath,
  getMyInsightsDriveItemById,
  getSiteDriveItemById,
  getSiteDriveItemByPath,
  getUserDriveItemById,
  getUserDriveItemByPath,
  getUserInsightsDriveItemById
} from '../../graph/graph.files';
import { formatBytes, getRelativeDisplayDate } from '../../utils/Utils';
import { OfficeGraphInsightString, ViewType } from '../../graph/types';
import { getFileTypeIconUriByExtension } from '../../styles/fluent-icons';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';
import { strings } from './strings';

/**
 * The File component is used to represent an individual file/folder from OneDrive or SharePoint by displaying information such as the file/folder name, an icon indicating the file type, and other properties such as the author, last modified date, or other details selected by the developer.
 *
 * @export
 * @class MgtFile
 * @extends {MgtTemplatedComponent}
 *
 * @cssprop --file-type-icon-height - {Length} file type icon height. Default value is 28px.
 * @cssprop --file-border - {String} file item border style. Default value is "1px solid transparent".
 * @cssprop --file-border-radius - {String} the border radius of the file component. Default value is 4px.
 * @cssprop --file-box-shadow - {String} the box-shadow of the component. Default value is none.
 * @cssprop --file-background-color - {Color} the background-color of the component.
 * @cssprop --file-background-color-focus - {Color} the background-color of the component on focus.
 * @cssprop --file-background-color-hover - {Color} the background-color of the component on hover.
 * @cssprop --file-padding - {String} the padding around the file component. Default value is 0px.
 * @cssprop --file-padding-inline-start - {Length} the padding between file icon and file details. Default value is 14px.
 * @cssprop --file-margin - {String} the margin around the file component. Default value is 0px.
 * @cssprop --file-line1-font-size - {Length} the first line text font size. Default value is 12px.
 * @cssprop --file-line1-font-weight - {Length} the first line text font weight. Default value is 400.
 * @cssprop --file-line1-color - {Color} the first line text color.
 * @cssprop --file-line1-text-transform - {String} the first line text text transform. Default value is initial.
 * @cssprop --file-line2-font-size - {Length} the second line text font size. Default value is 12px.
 * @cssprop --file-line2-font-weight - {Length} the second line text font weight. Default value is 400.
 * @cssprop --file-line2-color - {Color} the second line text color.
 * @cssprop --file-line2-text-transform - {String} the second line text text transform. Default value is initial.
 * @cssprop --file-line3-font-size - {Length} the third line text font size. Default value is 12px.
 * @cssprop --file-line3-font-weight - {Length} the third line text font weight. Default value is 400.
 * @cssprop --file-line3-color - {Color} the third line text color.
 * @cssprop --file-line3-text-transform - {String} the third line text text transform. Default value is 400.
 */

@customElement('file')
export class MgtFile extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }
  protected get strings() {
    return strings;
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
    void this.requestStateUpdate();
  }

  /**
   * allows developer to provide site id for a file
   *
   * @type {string}
   * @memberof MgtFile
   */
  @property({
    attribute: 'site-id'
  })
  public get siteId(): string {
    return this._siteId;
  }
  public set siteId(value: string) {
    if (value === this._siteId) {
      return;
    }

    this._siteId = value;
    void this.requestStateUpdate();
  }

  /**
   * allows developer to provide drive id for a file
   *
   * @type {string}
   * @memberof MgtFile
   */
  @property({
    attribute: 'drive-id'
  })
  public get driveId(): string {
    return this._driveId;
  }
  public set driveId(value: string) {
    if (value === this._driveId) {
      return;
    }

    this._driveId = value;
    void this.requestStateUpdate();
  }

  /**
   * allows developer to provide group id for a file
   *
   * @type {string}
   * @memberof MgtFile
   */
  @property({
    attribute: 'group-id'
  })
  public get groupId(): string {
    return this._groupId;
  }
  public set groupId(value: string) {
    if (value === this._groupId) {
      return;
    }

    this._groupId = value;
    void this.requestStateUpdate();
  }

  /**
   * allows developer to provide list id for a file
   *
   * @type {string}
   * @memberof MgtFile
   */
  @property({
    attribute: 'list-id'
  })
  public get listId(): string {
    return this._listId;
  }
  public set listId(value: string) {
    if (value === this._listId) {
      return;
    }

    this._listId = value;
    void this.requestStateUpdate();
  }

  /**
   * allows developer to provide user id for a file
   *
   * @type {string}
   * @memberof MgtFile
   */
  @property({
    attribute: 'user-id'
  })
  public get userId(): string {
    return this._userId;
  }
  public set userId(value: string) {
    if (value === this._userId) {
      return;
    }

    this._userId = value;
    void this.requestStateUpdate();
  }

  /**
   * allows developer to provide item id for a file
   *
   * @type {string}
   * @memberof MgtFile
   */
  @property({
    attribute: 'item-id'
  })
  public get itemId(): string {
    return this._itemId;
  }
  public set itemId(value: string) {
    if (value === this._itemId) {
      return;
    }

    this._itemId = value;
    void this.requestStateUpdate();
  }

  /**
   * allows developer to provide item path for a file
   *
   * @type {string}
   * @memberof MgtFile
   */
  @property({
    attribute: 'item-path'
  })
  public get itemPath(): string {
    return this._itemPath;
  }
  public set itemPath(value: string) {
    if (value === this._itemPath) {
      return;
    }

    this._itemPath = value;
    void this.requestStateUpdate();
  }

  /**
   * allows developer to provide insight type for a file
   * can be trending, used, or shared
   *
   * @type {OfficeGraphInsightString}
   * @memberof MgtFile
   */
  @property({
    attribute: 'insight-type'
  })
  public get insightType(): OfficeGraphInsightString {
    return this._insightType;
  }
  public set insightType(value: OfficeGraphInsightString) {
    if (value === this._insightType) {
      return;
    }

    this._insightType = value;
    void this.requestStateUpdate();
  }

  /**
   * allows developer to provide insight id for a file
   *
   * @type {string}
   * @memberof MgtFile
   */
  @property({
    attribute: 'insight-id'
  })
  public get insightId(): string {
    return this._insightId;
  }
  public set insightId(value: string) {
    if (value === this._insightId) {
      return;
    }

    this._insightId = value;
    void this.requestStateUpdate();
  }

  /**
   * allows developer to provide DriveItem object
   *
   * @type {MicrosoftGraph.DriveItem}
   * @memberof MgtFile
   */
  @property({
    type: Object
  })
  public get fileDetails(): DriveItem {
    return this._fileDetails;
  }
  public set fileDetails(value: DriveItem) {
    if (value === this._fileDetails) {
      return;
    }

    this._fileDetails = value;
    void this.requestStateUpdate();
  }

  /**
   * allows developer to provide file type icon url
   *
   * @type {string}
   * @memberof MgtFile
   */
  @property({
    attribute: 'file-icon'
  })
  public get fileIcon(): string {
    return this._fileIcon;
  }
  public set fileIcon(value: string) {
    if (value === this._fileIcon) {
      return;
    }

    this._fileIcon = value;
    void this.requestStateUpdate();
  }

  /**
   * object containing Graph details on item
   *
   * @type {MicrosoftGraph.DriveItem}
   * @memberof MgtFile
   */
  @property({ type: Object })
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
    attribute: 'view',
    converter: value => {
      if (!value || value.length === 0) {
        return ViewType.threelines;
      }

      value = value.toLowerCase();

      if (typeof ViewType[value] === 'undefined') {
        return ViewType.threelines;
      } else {
        return ViewType[value] as ViewType;
      }
    }
  })
  public view: ViewType;

  /**
   * Get the scopes required for file
   *
   * @static
   * @return {*}  {string[]}
   * @memberof MgtFile
   */
  public static get requiredScopes(): string[] {
    return [...new Set(['files.read', 'files.read.all', 'sites.read.all'])];
  }

  private _fileQuery: string;
  private _siteId: string;
  private _itemId: string;
  private _driveId: string;
  private _itemPath: string;
  private _listId: string;
  private _groupId: string;
  private _userId: string;
  private _insightType: OfficeGraphInsightString;
  private _insightId: string;
  private _fileDetails: DriveItem;
  private _fileIcon: string;

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
    let fileTemplate: TemplateResult;

    fileTemplate = this.renderTemplate('default', { file });
    if (!fileTemplate) {
      const fileDetailsTemplate: TemplateResult = this.renderDetails(file);
      const fileTypeIconTemplate: TemplateResult = this.renderFileTypeIcon();

      fileTemplate = html`
        <div class="item">
          ${fileTypeIconTemplate} ${fileDetailsTemplate}
        </div>`;
    }

    return fileTemplate;
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
    if (!this.fileIcon && !this.driveItem.name) {
      return html``;
    }

    let fileIconSrc;

    if (this.fileIcon) {
      fileIconSrc = this.fileIcon;
    } else {
      // get file type extension from file name
      const re = /(?:\.([^.]+))?$/;
      const fileType =
        this.driveItem.package === undefined && this.driveItem.folder === undefined
          ? re.exec(this.driveItem.name)[1]
            ? re.exec(this.driveItem.name)[1].toLowerCase()
            : this.driveItem.size
            ? 'null'
            : 'folder'
          : this.driveItem.package !== undefined
          ? this.driveItem.package.type === 'oneNote'
            ? 'onetoc'
            : 'folder'
          : 'folder';
      fileIconSrc = getFileTypeIconUriByExtension(fileType, 48, 'svg');
    }

    return html`
      <div class="item__file-type-icon">
        ${
          fileIconSrc
            ? html`
              <img src=${fileIconSrc} alt="File icon" />
            `
            : html`
              ${getSvg(SvgIcon.File)}
            `
        }
      </div>
    `;
  }

  /**
   * Render the file details
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

  /**
   * load state into the component.
   *
   * @protected
   * @returns
   * @memberof MgtFile
   */
  protected async loadState() {
    if (this.fileDetails) {
      this.driveItem = this.fileDetails;
      return;
    }

    const provider = Providers.globalProvider;
    if (!provider || provider.state === ProviderState.Loading) {
      return;
    }

    if (provider.state === ProviderState.SignedOut) {
      this.driveItem = null;
      return;
    }

    const graph = provider.graph.forComponent(this);
    let driveItem: DriveItem;

    // evaluate to true when only item-id or item-path is provided
    const getFromMyDrive = !this.driveId && !this.siteId && !this.groupId && !this.listId && !this.userId;

    if (
      // return null when a combination of provided properties are required
      (this.driveId && !this.itemId && !this.itemPath) ||
      (this.siteId && !this.itemId && !this.itemPath) ||
      (this.groupId && !this.itemId && !this.itemPath) ||
      (this.listId && !this.siteId && !this.itemId) ||
      (this.insightType && !this.insightId) ||
      (this.userId && !this.itemId && !this.itemPath && !this.insightType && !this.insightId)
    ) {
      driveItem = null;
    } else if (this.fileQuery) {
      driveItem = await getDriveItemByQuery(graph, this.fileQuery);
    } else if (this.itemId && getFromMyDrive) {
      driveItem = await getMyDriveItemById(graph, this.itemId);
    } else if (this.itemPath && getFromMyDrive) {
      driveItem = await getMyDriveItemByPath(graph, this.itemPath);
    } else if (this.userId) {
      if (this.itemId) {
        driveItem = await getUserDriveItemById(graph, this.userId, this.itemId);
      } else if (this.itemPath) {
        driveItem = await getUserDriveItemByPath(graph, this.userId, this.itemPath);
      } else if (this.insightType && this.insightId) {
        driveItem = await getUserInsightsDriveItemById(graph, this.userId, this.insightType, this.insightId);
      }
    } else if (this.driveId) {
      if (this.itemId) {
        driveItem = await getDriveItemById(graph, this.driveId, this.itemId);
      } else if (this.itemPath) {
        driveItem = await getDriveItemByPath(graph, this.driveId, this.itemPath);
      }
    } else if (this.siteId && !this.listId) {
      if (this.itemId) {
        driveItem = await getSiteDriveItemById(graph, this.siteId, this.itemId);
      } else if (this.itemPath) {
        driveItem = await getSiteDriveItemByPath(graph, this.siteId, this.itemPath);
      }
    } else if (this.listId) {
      driveItem = await getListDriveItemById(graph, this.siteId, this.listId, this.itemId);
    } else if (this.groupId) {
      if (this.itemId) {
        driveItem = await getGroupDriveItemById(graph, this.groupId, this.itemId);
      } else if (this.itemPath) {
        driveItem = await getGroupDriveItemByPath(graph, this.groupId, this.itemPath);
      }
    } else if (this.insightType && !this.userId) {
      driveItem = await getMyInsightsDriveItemById(graph, this.insightType, this.insightId);
    }

    this.driveItem = driveItem;
  }

  private getTextFromProperty(driveItem: DriveItem, properties: string): string {
    if (!properties || properties.length === 0) {
      return null;
    }

    const propertyList = properties.trim().split(',');
    let text: string;
    let i = 0;

    while (!text && i < propertyList.length) {
      const current = propertyList[i].trim();
      switch (current) {
        case 'size': {
          // convert size to kb, mb, gb
          let size = '0';
          if (driveItem.size) {
            size = formatBytes(driveItem.size);
          }
          text = `${this.strings.sizeSubtitle}: ${size}`;
          break;
        }
        case 'lastModifiedDateTime': {
          // convert date time
          let relativeDateString: string;
          let lastModifiedString: string;
          if (driveItem.lastModifiedDateTime) {
            const lastModifiedDateTime = new Date(driveItem.lastModifiedDateTime);
            relativeDateString = getRelativeDisplayDate(lastModifiedDateTime);
            lastModifiedString = `${this.strings.modifiedSubtitle} ${relativeDateString}`;
          } else {
            lastModifiedString = '';
          }
          text = lastModifiedString;
          break;
        }
        default:
          text = driveItem[current] as string;
      }
      i++;
    }

    return text;
  }
}
