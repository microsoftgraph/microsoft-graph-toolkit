/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  arraysAreEqual,
  GraphPageIterator,
  Providers,
  ProviderState,
  customElement,
  mgtHtml,
  MgtTemplatedComponent
} from '@microsoft/mgt-element';
import { DriveItem } from '@microsoft/microsoft-graph-types';
import { html, TemplateResult } from 'lit';
import { property, state } from 'lit/decorators.js';
import { repeat } from 'lit/directives/repeat.js';
import {
  clearFilesCache,
  fetchNextAndCacheForFilesPageIterator,
  getDriveFilesByIdIterator,
  getDriveFilesByPathIterator,
  getFilesByIdIterator,
  getFilesByListQueryIterator,
  getFilesByPathIterator,
  getFilesByQueries,
  getFilesIterator,
  getGroupFilesByIdIterator,
  getGroupFilesByPathIterator,
  getMyInsightsFiles,
  getSiteFilesByIdIterator,
  getSiteFilesByPathIterator,
  getUserFilesByIdIterator,
  getUserFilesByPathIterator,
  getUserInsightsFiles
} from '../../graph/graph.files';
import './mgt-file-upload/mgt-file-upload';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';
import { OfficeGraphInsightString, ViewType } from '../../graph/types';
import { styles } from './mgt-file-list-css';
import { strings } from './strings';
import { MgtFile } from '../mgt-file/mgt-file';
import { MgtFileUploadConfig } from './mgt-file-upload/mgt-file-upload';

import { fluentProgressRing } from '@fluentui/web-components';
import { registerFluentComponents } from '../../utils/FluentComponents';
import { CardSection } from '../BasePersonCardSection';

registerFluentComponents(fluentProgressRing);

/**
 * The File List component displays a list of multiple folders and files by
 * using the file/folder name, an icon, and other properties specified by the developer.
 * This component uses the mgt-file component.
 *
 * @export
 * @class MgtFileList
 *
 * @fires {CustomEvent<MicrosoftGraph.DriveItem>} itemClick - Fired when a user clicks on a file.
 * it returns the file (DriveItem) details.
 *
 * NOTE: This component also allows customizing the tokens from mgt-file and mgt-file-upload components.
 * @cssprop --file-list-background-color - {Color} the background color of the component.
 * @cssprop --file-list-box-shadow - {String} the box-shadow syle of the component. Default value is --elevation-shadow-card-rest.
 * @cssprop --file-list-border-radius - {Length} the file list box border radius. Default value is 8px.
 * @cssprop --file-list-border - {String} the file list border style. Default value is none.
 * @cssprop --file-list-padding -{String} the file list padding.  Default value is 0px.
 * @cssprop --file-list-margin -{String} the file list margin. Default value is 0px.
 * @cssprop --show-more-button-background-color - {Color} the "show more" button background color.
 * @cssprop --show-more-button-background-color--hover - {Color} the "show more" button background color on hover.
 * @cssprop --show-more-button-font-size - {String} the "show more" text font size. Default value is 12px.
 * @cssprop --show-more-button-padding - {String} the "show more" button padding. Default value is 0px.
 * @cssprop --show-more-button-border-bottom-right-radius - {String} the "show more" button bottom right border radius. Default value is 8px.
 * @cssprop --show-more-button-border-bottom-left-radius - {String} the "show more" button bottom left border radius. Default value is 8px;
 * @cssprop --progress-ring-size -{String} Progress ring height and width. Default value is 24px.
 */

@customElement('file-list')
export class MgtFileList extends MgtTemplatedComponent implements CardSection {
  @state() private _isCompact = false;
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  protected get strings(): Record<string, string> {
    return strings;
  }

  /**
   * allows developer to provide query for a file list
   *
   * @type {string}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'file-list-query'
  })
  public get fileListQuery(): string {
    return this._fileListQuery;
  }
  public set fileListQuery(value: string) {
    if (value === this._fileListQuery) {
      return;
    }

    this._fileListQuery = value;
    void this.requestStateUpdate(true);
  }

  /**
   * The name for display in the overview section.
   *
   * @readonly
   * @type {string}
   * @memberof MgtFileList
   */
  public get displayName(): string {
    return this.strings.filesSectionTitle;
  }

  /**
   * The title for the card when rendered as a card full.
   *
   * @readonly
   * @type {string}
   * @memberof MgtFileList
   */
  public get cardTitle(): string {
    return this.strings.filesSectionTitle;
  }

  /**
   * Render the icon for display in the navigation ribbon.
   *
   * @returns {TemplateResult}
   * @memberof MgtFileList
   */
  public renderIcon(): TemplateResult {
    return getSvg(SvgIcon.Files);
  }

  /**
   * allows developer to provide an array of file queries
   *
   * @type {string[]}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'file-queries',
    converter: (value, _type) => {
      if (value) {
        return value.split(',').map(v => v.trim());
      } else {
        return null;
      }
    }
  })
  public get fileQueries(): string[] {
    return this._fileQueries;
  }
  public set fileQueries(value: string[]) {
    if (arraysAreEqual(this._fileQueries, value)) {
      return;
    }

    this._fileQueries = value;
    void this.requestStateUpdate(true);
  }

  /**
   * allows developer to provide an array of files
   *
   * @type {MicrosoftGraph.DriveItem[]}
   * @memberof MgtFileList
   */
  @property({ type: Object })
  public files: DriveItem[];

  /**
   * allows developer to provide site id for a file
   *
   * @type {string}
   * @memberof MgtFileList
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
    void this.requestStateUpdate(true);
  }

  /**
   * allows developer to provide drive id for a file
   *
   * @type {string}
   * @memberof MgtFileList
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
    void this.requestStateUpdate(true);
  }

  /**
   * allows developer to provide group id for a file
   *
   * @type {string}
   * @memberof MgtFileList
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
    void this.requestStateUpdate(true);
  }

  /**
   * allows developer to provide item id for a file
   *
   * @type {string}
   * @memberof MgtFileList
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
    void this.requestStateUpdate(true);
  }

  /**
   * allows developer to provide item path for a file
   *
   * @type {string}
   * @memberof MgtFileList
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
    void this.requestStateUpdate(true);
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
    void this.requestStateUpdate(true);
  }

  /**
   * allows developer to provide insight type for a file
   * can be trending, used, or shared
   *
   * @type {OfficeGraphInsightString}
   * @memberof MgtFileList
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
    void this.requestStateUpdate(true);
  }

  /**
   * Sets what data to be rendered (file icon only, oneLine, twoLines threeLines).
   * Default is 'threeLines'.
   *
   * @type {ViewType}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'item-view',
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
  public itemView: ViewType;

  /**
   * allows developer to provide file type to filter the list
   * can be docx
   *
   * @type {string[]}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'file-extensions',
    converter: (value, _type) => {
      return value.split(',').map(v => v.trim());
    }
  })
  public get fileExtensions(): string[] {
    return this._fileExtensions;
  }
  public set fileExtensions(value: string[]) {
    if (arraysAreEqual(this._fileExtensions, value)) {
      return;
    }

    this._fileExtensions = value;
    void this.requestStateUpdate(true);
  }

  /**
   * A number value to indicate the number of more files to load when show more button is clicked
   *
   * @type {number}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'page-size',
    type: Number
  })
  public get pageSize(): number {
    return this._pageSize;
  }
  public set pageSize(value: number) {
    if (value === this._pageSize) {
      return;
    }

    this._pageSize = value;
    void this.requestStateUpdate(true);
  }

  @property({
    attribute: 'disable-open-on-click',
    type: Boolean
  })
  public disableOpenOnClick = false;
  /**
   * A boolean value indication if 'show-more' button should be disabled
   *
   * @type {boolean}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'hide-more-files-button',
    type: Boolean
  })
  public hideMoreFilesButton: boolean;

  /**
   * A number value indication for file size upload (KB)
   *
   * @type {number}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'max-file-size',
    type: Number
  })
  public get maxFileSize(): number {
    return this._maxFileSize;
  }
  public set maxFileSize(value: number) {
    if (value === this._maxFileSize) {
      return;
    }

    this._maxFileSize = value;
    void this.requestStateUpdate(true);
  }

  /**
   * A boolean value indication if file upload extension should be enable or disabled
   *
   * @type {boolean}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'enable-file-upload',
    type: Boolean
  })
  public enableFileUpload: boolean;

  /**
   * A number value to indicate the max number allowed of files to upload.
   *
   * @type {number}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'max-upload-file',
    type: Number
  })
  public get maxUploadFile(): number {
    return this._maxUploadFile;
  }
  public set maxUploadFile(value: number) {
    if (value === this._maxUploadFile) {
      return;
    }

    this._maxUploadFile = value;
    void this.requestStateUpdate(true);
  }

  /**
   * A Array of file extensions to be excluded from file upload.
   *
   * @type {string[]}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'excluded-file-extensions',
    converter: (value, _type) => {
      return value.split(',').map(v => v.trim());
    }
  })
  public get excludedFileExtensions(): string[] {
    return this._excludedFileExtensions;
  }
  public set excludedFileExtensions(value: string[]) {
    if (arraysAreEqual(this._excludedFileExtensions, value)) {
      return;
    }

    this._excludedFileExtensions = value;
    void this.requestStateUpdate(true);
  }

  /**
   * Get the scopes required for file list
   *
   * @static
   * @return {*}  {string[]}
   * @memberof MgtFileList
   */
  public static get requiredScopes(): string[] {
    return [...new Set([...MgtFile.requiredScopes])];
  }

  private _fileListQuery: string;
  private _fileQueries: string[];
  private _siteId: string;
  private _itemId: string;
  private _driveId: string;
  private _itemPath: string;
  private _groupId: string;
  private _insightType: OfficeGraphInsightString;
  private _fileExtensions: string[];
  private _pageSize: number;
  private _excludedFileExtensions: string[];
  private _maxUploadFile: number;
  private _maxFileSize: number;
  private _userId: string;
  private _preloadedFiles: DriveItem[];
  private pageIterator: GraphPageIterator<DriveItem>;
  // tracking user arrow key input of selection for accessibility purpose
  private _focusedItemIndex = -1;

  @state() private _isLoadingMore: boolean;

  constructor() {
    super();

    this.pageSize = 10;
    this.itemView = ViewType.twolines;
    this.maxUploadFile = 10;
    this.enableFileUpload = false;
    this._preloadedFiles = [];
  }

  /**
   * Override requestStateUpdate to include clearstate.
   *
   * @memberof MgtFileList
   */
  protected requestStateUpdate(force?: boolean) {
    this.clearState();
    return super.requestStateUpdate(force);
  }

  /**
   * Reset state
   *
   * @memberof MgtFileList
   */
  protected clearState(): void {
    super.clearState();
    this.files = null;
  }

  /**
   * Set the section to compact view mode
   *
   * @returns
   * @memberof BasePersonCardSection
   */
  public asCompactView() {
    this._isCompact = true;
    return this;
  }

  /**
   * Set the section to full view mode
   *
   * @returns
   * @memberof BasePersonCardSection
   */
  public asFullView() {
    this._isCompact = false;
    return this;
  }

  /**
   * Render the file list
   *
   * @return {*}
   * @memberof MgtFileList
   */
  public render() {
    if (!this.files && this.isLoadingState) {
      return this.renderLoading();
    }

    if (!this.files || this.files.length === 0) {
      return this.renderNoData();
    }

    return this._isCompact ? this.renderCompactView() : this.renderFullView();
  }

  /**
   * Render the compact view
   *
   * @returns {TemplateResult}
   * @memberof MgtFileList
   */
  public renderCompactView(): TemplateResult {
    const files = this.files.slice(0, 3);

    return this.renderFiles(files);
  }

  /**
   * Render the full view
   *
   * @returns {TemplateResult}
   * @memberof MgtFileList
   */
  public renderFullView(): TemplateResult {
    return this.renderTemplate('default', { files: this.files }) || this.renderFiles(this.files);
  }

  /**
   * Render the loading state
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtFileList
   */
  protected renderLoading(): TemplateResult {
    return this.renderTemplate('loading', null) || html``;
  }

  /**
   * Render the state when no data is available
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtFileList
   */
  protected renderNoData(): TemplateResult {
    return (
      this.renderTemplate('no-data', null) ||
      (this.enableFileUpload === true && Providers.globalProvider !== undefined
        ? html`
            <div class="file-list-wrapper" dir=${this.direction}>
              ${this.renderFileUpload()}
            </div>`
        : html``)
    );
  }

  /**
   * Render the list of files.
   *
   * @protected
   * @param {*} files
   * @returns {TemplateResult}
   * @memberof mgtFileList
   */
  protected renderFiles(files: DriveItem[]): TemplateResult {
    return html`
      <div id="file-list-wrapper" class="file-list-wrapper" dir=${this.direction}>
        ${this.enableFileUpload ? this.renderFileUpload() : null}
        <ul
          id="file-list"
          class="file-list"
        >
          <li
            tabindex="0"
            class="file-item"
            @keydown="${this.onFileListKeyDown}"
            @focus="${this.onFocusFirstItem}"
            @click=${(e: UIEvent) => this.handleItemSelect(files[0], e)}>
            ${this.renderFile(files[0])}
          </li>
          ${repeat(
            files.slice(1),
            f => f.id,
            f => html`
              <li
                class="file-item"
                @keydown="${this.onFileListKeyDown}"
                @click=${(e: UIEvent) => this.handleItemSelect(f, e)}>
                ${this.renderFile(f)}
              </li>
            `
          )}
        </ul>
        ${
          !this.hideMoreFilesButton &&
          this.pageIterator &&
          (this.pageIterator.hasNext || this._preloadedFiles.length) &&
          !this._isCompact
            ? this.renderMoreFileButton()
            : null
        }
      </div>
    `;
  }

  /**
   * Render an individual file.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof mgtFileList
   */
  protected renderFile(file: DriveItem): TemplateResult {
    const view = this.itemView;
    return (
      this.renderTemplate('file', { file }, file.id) ||
      mgtHtml`
        <mgt-file class="mgt-file-item" .fileDetails=${file} .view=${view}></mgt-file>
      `
    );
  }

  /**
   * Render the button when clicked will show more files.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtFileList
   */
  protected renderMoreFileButton(): TemplateResult {
    if (this._isLoadingMore) {
      return html`
        <fluent-progress-ring role="progressbar" viewBox="0 0 8 8" class="progress-ring"></fluent-progress-ring>
      `;
    } else {
      return html`
        <fluent-button
          appearance="stealth"
          id="show-more"
          class="show-more"
          @click=${() => this.renderNextPage()}
        >
          <span class="show-more-text">${this.strings.showMoreSubtitle}</span>
        </fluent-button>`;
    }
  }

  /**
   * Render MgtFileUpload sub component
   *
   * @returns
   */
  protected renderFileUpload(): TemplateResult {
    const fileUploadConfig: MgtFileUploadConfig = {
      graph: Providers.globalProvider.graph.forComponent(this),
      driveId: this.driveId,
      excludedFileExtensions: this.excludedFileExtensions,
      groupId: this.groupId,
      itemId: this.itemId,
      itemPath: this.itemPath,
      userId: this.userId,
      siteId: this.siteId,
      maxFileSize: this.maxFileSize,
      maxUploadFile: this.maxUploadFile
    };
    return mgtHtml`
        <mgt-file-upload .fileUploadList=${fileUploadConfig} ></mgt-file-upload>
      `;
  }

  /**
   * Handles setting the focusedItemIndex to 0 when you focus on the first item
   * in the file list.
   *
   * @returns void
   */
  private readonly onFocusFirstItem = () => (this._focusedItemIndex = 0);

  /**
   * Handle accessibility keyboard keydown events (arrow up, arrow down, enter, tab) on file list
   *
   * @param event
   */
  private readonly onFileListKeyDown = (event: KeyboardEvent): void => {
    const fileList = this.renderRoot.querySelector('.file-list');
    let focusedItem: HTMLElement;

    if (!fileList?.children.length) {
      return;
    }

    if (event.code === 'ArrowUp' || event.code === 'ArrowDown') {
      if (event.code === 'ArrowUp') {
        if (this._focusedItemIndex === -1) {
          this._focusedItemIndex = fileList.children.length;
        }
        this._focusedItemIndex = (this._focusedItemIndex - 1 + fileList.children.length) % fileList.children.length;
      }
      if (event.code === 'ArrowDown') {
        this._focusedItemIndex = (this._focusedItemIndex + 1) % fileList.children.length;
      }

      focusedItem = fileList.children[this._focusedItemIndex] as HTMLElement;
      this.updateItemBackgroundColor(fileList, focusedItem, 'focused');
    }

    if (event.code === 'Enter' || event.code === 'Space') {
      focusedItem = fileList.children[this._focusedItemIndex] as HTMLElement;

      const file = focusedItem.children[0] as MgtFile;
      event.preventDefault();
      this.fireCustomEvent('itemClick', file.fileDetails);
      this.handleFileClick(file.fileDetails);

      this.updateItemBackgroundColor(fileList, focusedItem, 'selected');
    }

    if (event.code === 'Tab') {
      focusedItem = fileList.children[this._focusedItemIndex] as HTMLElement;
    }
  };

  /**
   * load state into the component.
   *
   * @protected
   * @returns
   * @memberof MgtFileList
   */
  protected async loadState() {
    const provider = Providers.globalProvider;
    if (!provider || provider.state === ProviderState.Loading) {
      return;
    }

    if (provider.state === ProviderState.SignedOut) {
      this.files = null;
      return;
    }
    const graph = provider.graph.forComponent(this);
    let files: DriveItem[];
    let pageIterator: GraphPageIterator<DriveItem>;

    const getFromMyDrive = !this.driveId && !this.siteId && !this.groupId && !this.userId;

    // combinations of these attributes must be provided in order for the component to know which endpoint to call to request files
    // not supplying enough for these combinations will get a null file result
    if (
      (this.driveId && !this.itemId && !this.itemPath) ||
      (this.groupId && !this.itemId && !this.itemPath) ||
      (this.siteId && !this.itemId && !this.itemPath) ||
      (this.userId && !this.insightType && !this.itemId && !this.itemPath)
    ) {
      this.files = null;
    }

    if (!this.files) {
      if (this.fileListQuery) {
        pageIterator = await getFilesByListQueryIterator(graph, this.fileListQuery, this.pageSize);
      } else if (this.fileQueries) {
        files = await getFilesByQueries(graph, this.fileQueries);
      } else if (getFromMyDrive) {
        if (this.itemId) {
          pageIterator = await getFilesByIdIterator(graph, this.itemId, this.pageSize);
        } else if (this.itemPath) {
          pageIterator = await getFilesByPathIterator(graph, this.itemPath, this.pageSize);
        } else if (this.insightType) {
          files = await getMyInsightsFiles(graph, this.insightType);
        } else {
          pageIterator = await getFilesIterator(graph, this.pageSize);
        }
      } else if (this.driveId) {
        if (this.itemId) {
          pageIterator = await getDriveFilesByIdIterator(graph, this.driveId, this.itemId, this.pageSize);
        } else if (this.itemPath) {
          pageIterator = await getDriveFilesByPathIterator(graph, this.driveId, this.itemPath, this.pageSize);
        }
      } else if (this.groupId) {
        if (this.itemId) {
          pageIterator = await getGroupFilesByIdIterator(graph, this.groupId, this.itemId, this.pageSize);
        } else if (this.itemPath) {
          pageIterator = await getGroupFilesByPathIterator(graph, this.groupId, this.itemPath, this.pageSize);
        }
      } else if (this.siteId) {
        if (this.itemId) {
          pageIterator = await getSiteFilesByIdIterator(graph, this.siteId, this.itemId, this.pageSize);
        } else if (this.itemPath) {
          pageIterator = await getSiteFilesByPathIterator(graph, this.siteId, this.itemPath, this.pageSize);
        }
      } else if (this.userId) {
        if (this.itemId) {
          pageIterator = await getUserFilesByIdIterator(graph, this.userId, this.itemId, this.pageSize);
        } else if (this.itemPath) {
          pageIterator = await getUserFilesByPathIterator(graph, this.userId, this.itemPath, this.pageSize);
        } else if (this.insightType) {
          files = await getUserInsightsFiles(graph, this.userId, this.insightType);
        }
      }

      if (pageIterator) {
        this.pageIterator = pageIterator;
        this._preloadedFiles = [...this.pageIterator.value];

        // handle when cached file length is greater than page size
        if (this._preloadedFiles.length >= this.pageSize) {
          files = this._preloadedFiles.splice(0, this.pageSize);
        } else {
          files = this._preloadedFiles.splice(0, this._preloadedFiles.length);
        }
      }

      // filter files when extensions are provided
      let filteredByFileExtension: DriveItem[];
      if (this.fileExtensions && this.fileExtensions !== null) {
        // retrive all pages before filtering
        if (this.pageIterator?.value) {
          while (this.pageIterator.hasNext) {
            await fetchNextAndCacheForFilesPageIterator(this.pageIterator);
          }
          files = this.pageIterator.value;
          this._preloadedFiles = [];
        }
        filteredByFileExtension = files.filter(file => {
          for (const e of this.fileExtensions) {
            if (e === this.getFileExtension(file.name)) {
              return file;
            }
          }
        });
      }

      if (filteredByFileExtension?.length >= 0) {
        this.files = filteredByFileExtension;
        if (this.pageSize) {
          files = this.files.splice(0, this.pageSize);
          this.files = files;
        }
      } else {
        this.files = files;
      }
    }
  }

  /**
   * Handle the click event on an item.
   *
   * @protected
   * @memberof MgtFileList
   */
  protected handleItemSelect(item: DriveItem, event: UIEvent): void {
    this.handleFileClick(item);
    this.fireCustomEvent('itemClick', item);

    // handle accessibility updates when item clicked
    if (event) {
      const fileList = this.renderRoot.querySelector('.file-list');

      // get index of the focused item
      const nodes = Array.from(fileList.children);
      const li = (event.target as HTMLElement).closest('li');
      const index = nodes.indexOf(li);
      this._focusedItemIndex = index;
      const clickedItem = fileList.children[this._focusedItemIndex] as HTMLElement;
      this.updateItemBackgroundColor(fileList, clickedItem, 'selected');
    }
  }

  /**
   * Handle the click event on button to show next page.
   *
   * @protected
   * @memberof MgtFileList
   */
  protected async renderNextPage() {
    // render next page from cache if exists, or else use iterator
    if (this._preloadedFiles.length > 0) {
      this.files = [
        ...this.files,
        ...this._preloadedFiles.splice(0, Math.min(this.pageSize, this._preloadedFiles.length))
      ];
    } else {
      if (this.pageIterator.hasNext) {
        this._isLoadingMore = true;
        const root = this.renderRoot.querySelector('#file-list-wrapper');
        if (root?.animate) {
          // play back
          root.animate(
            [
              {
                height: 'auto',
                transformOrigin: 'top left'
              },
              {
                height: 'auto',
                transformOrigin: 'top left'
              }
            ],
            {
              duration: 1000,
              easing: 'ease-in-out',
              fill: 'both'
            }
          );
        }
        await fetchNextAndCacheForFilesPageIterator(this.pageIterator);
        this._isLoadingMore = false;
        this.files = this.pageIterator.value;
      }
    }

    this.requestUpdate();
  }

  private handleFileClick(file: DriveItem) {
    if (file?.webUrl && !this.disableOpenOnClick) {
      window.open(file.webUrl, '_blank', 'noreferrer');
    }
  }

  /**
   * Get file extension string from file name
   *
   * @param name file name
   * @returns {string} file extension
   */
  private getFileExtension(name: string) {
    const re = /(?:\.([^.]+))?$/;
    const fileExtension = re.exec(name)[1] || '';

    return fileExtension;
  }

  /**
   * Handle remove and add css class on accessibility keyboard select and focus
   *
   * @param fileList HTML element
   * @param focusedItem HTML element
   * @param className background class to be applied
   */
  private updateItemBackgroundColor(fileList: Element, focusedItem: HTMLElement, className: string) {
    // reset background color and remove tabindex
    for (const node of fileList.children) {
      node.classList.remove(className);
      node.removeAttribute('tabindex');
    }

    // set focused item background color
    if (focusedItem) {
      focusedItem.classList.add(className);
      focusedItem.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'start' });
      focusedItem.setAttribute('tabindex', '0');
      focusedItem.focus();
    }

    // remove selected classes
    for (const node of fileList.children) {
      node.classList.remove('selected');
    }
  }

  /**
   * Handle reload of File List and condition to clear cache
   *
   * @param clearCache boolean, if true clear cache
   */
  public reload(clearCache = false) {
    if (clearCache) {
      // clear cache File List
      void clearFilesCache();
    }

    void this.requestStateUpdate(true);
  }
}
