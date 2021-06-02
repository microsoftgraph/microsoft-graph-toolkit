/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  arraysAreEqual,
  GraphPageIterator,
  MgtTemplatedComponent,
  Providers,
  ProviderState
} from '@microsoft/mgt-element';
import { DriveItem } from '@microsoft/microsoft-graph-types';
import { customElement, html, property, TemplateResult } from 'lit-element';
import { repeat } from 'lit-html/directives/repeat';
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
import '../sub-components/mgt-spinner/mgt-spinner';
import { OfficeGraphInsightString, ViewType } from '../../graph/types';
import { styles } from './mgt-file-list-css';
import { strings } from './strings';
import { MgtFile } from '../mgt-file/mgt-file';

/**
 * The File List component displays a list of multiple folders and files by
 * using the file/folder name, an icon, and other properties specicified by the developer.
 * This component uses the mgt-file component.
 *
 * @export
 * @class MgtFileList
 * @extends {MgtTemplatedComponent}
 *
 * @fires itemClick - Fired when user click a file. Returns the file (DriveItem) details.
 * @cssprop --file-list-background-color - {Color} File list background color
 * @cssprop --file-list-box-shadow - {String} File list box shadow style
 * @cssprop --file-list-border - {String} File list border styles
 * @cssprop --file-list-padding -{String} File list padding
 * @cssprop --file-list-margin -{String} File list margin
 * @cssprop --file-item-background-color--hover - {Color} File item background hover color
 * @cssprop --file-item-border-top - {String} File item border top style
 * @cssprop --file-item-border-left - {String} File item border left style
 * @cssprop --file-item-border-right - {String} File item border right style
 * @cssprop --file-item-border-bottom - {String} File item border bottom style
 * @cssprop --file-item-background-color--active - {Color} File item background active color
 * @cssprop --file-item-border-radius - {String} File item border radius
 * @cssprop --file-item-margin - {String} File item margin
 * @cssprop --show-more-button-background-color - {Color} Show more button background color
 * @cssprop --show-more-button-background-color--hover - {Color} Show more button background hover color
 * @cssprop --show-more-button-font-size - {String} Show more button font size
 * @cssprop --show-more-button-padding - {String} Show more button padding
 * @cssprop --show-more-button-border-bottom-right-radius - {String} Show more button bottom right radius
 * @cssprop --show-more-button-border-bottom-left-radius - {String} Show more button bottom left radius
 */

@customElement('mgt-file-list')
export class MgtFileList extends MgtTemplatedComponent {
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
    this.requestStateUpdate(true);
  }

  /**
   * allows developer to provide an array of file queries
   *
   * @type {string[]}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'file-queries',
    converter: (value, type) => {
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
    this.requestStateUpdate(true);
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
    this.requestStateUpdate(true);
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
    this.requestStateUpdate(true);
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
    this.requestStateUpdate(true);
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
    this.requestStateUpdate(true);
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
    this.requestStateUpdate(true);
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
    this.requestStateUpdate(true);
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
    this.requestStateUpdate(true);
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
        return ViewType[value];
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
    converter: (value, type) => {
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
    this.requestStateUpdate(true);
  }

  /**
   * A number value to indicate the number of more files to load when show more button is clicked
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
    this.requestStateUpdate(true);
  }

  /**
   * A boolean value indication if 'show-more' button should be disabled
   * @type {boolean}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'hide-more-files-button',
    type: Boolean
  })
  public hideMoreFilesButton: boolean;

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
  private _userId: string;
  private _preloadedFiles: DriveItem[];
  private pageIterator: GraphPageIterator<DriveItem>;
  // tracking user arrow key input of selection for accessibility purpose
  private _focusedItemIndex: number = -1;

  constructor() {
    super();

    this.pageSize = 10;
    this.itemView = ViewType.twolines;
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

  public render() {
    if (!this.files && this.isLoadingState) {
      return this.renderLoading();
    }

    if (!this.files || this.files.length === 0) {
      return this.renderNoData();
    }

    return this.renderTemplate('default', { files: this.files }) || this.renderFiles();
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
    return this.renderTemplate('no-data', null) || html``;
  }

  /**
   * Render the list of files.
   *
   * @protected
   * @param {*} files
   * @returns {TemplateResult}
   * @memberof mgtFileList
   */
  protected renderFiles(): TemplateResult {
    return html`
      <div id="file-list-wrapper" class="file-list-wrapper" dir=${this.direction}>
        <ul
          id="file-list"
          class="file-list"
          tabindex="0"
          @keydown="${this.onFileListKeyDown}"
          @keyup="${this.onFileListKeyUp}"
          @blur="${this.onFileListOut}"
        >
          ${repeat(
            this.files,
            f => f.id,
            f => html`
              <li class="file-item">
                ${this.renderFile(f)}
              </li>
            `
          )}
        </ul>
        ${
          !this.hideMoreFilesButton && this.pageIterator && (this.pageIterator.hasNext || this._preloadedFiles.length)
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
      html`
        <mgt-file .fileDetails=${file} .view=${view} @click=${e => this.handleItemSelect(file, e)}></mgt-file>
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
    if (this.isLoadingState) {
      return html`
        <mgt-spinner></mgt-spinner>
      `;
    } else {
      return html`<a id="show-more" class="show-more" @click=${() => this.renderNextPage()} tabindex="0" @keydown=${
        this.onShowMoreKeyDown
      }><span>${this.strings.showMoreSubtitle}<span></a>`;
    }
  }

  /**
   * Handle accessibility keyboard enter event on 'show more items' button
   *
   * @param event
   */
  private onShowMoreKeyDown(event: KeyboardEvent): void {
    if (event && event.code === 'Enter') {
      event.preventDefault();
      this.renderNextPage();
    }
  }

  /**
   * Handle accessibility keyboard keyup events on file list
   *
   * @param event
   */
  private onFileListKeyUp(event: KeyboardEvent): void {
    const fileList = this.renderRoot.querySelector('.file-list');
    const focusedItem = fileList.children[this._focusedItemIndex];

    if (event.code === 'Enter' || event.code === 'Space') {
      event.preventDefault();

      focusedItem?.classList.remove('selected');
      focusedItem?.classList.add('focused');
    }
  }

  /**
   * Handle accessibility keyboard keydown events (arrow up, arrow down, enter, tab) on file list
   *
   * @param event
   */
  private onFileListKeyDown(event: KeyboardEvent): void {
    const fileList = this.renderRoot.querySelector('.file-list');
    let focusedItem;

    if (!fileList || !fileList.children.length) {
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

      focusedItem = fileList.children[this._focusedItemIndex];
      this.updateItemBackgroundColor(fileList, focusedItem, 'focused');
    }

    if (event.code === 'Enter' || event.code === 'Space') {
      focusedItem = fileList.children[this._focusedItemIndex];

      const file = focusedItem.children[0] as any;
      event.preventDefault();
      this.fireCustomEvent('itemClick', file.fileDetails);

      this.updateItemBackgroundColor(fileList, focusedItem, 'selected');
    }

    if (event.code === 'Tab') {
      focusedItem = fileList.children[this._focusedItemIndex];
      focusedItem?.classList.remove('focused');
    }
  }

  /**
   * Remove accessibility keyboard focused when out of file list
   *
   */
  private onFileListOut() {
    const fileList = this.renderRoot.querySelector('.file-list');
    const focusedItem = fileList.children[this._focusedItemIndex];
    focusedItem?.classList.remove('focused');
  }

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
        if (this.pageIterator && this.pageIterator.value) {
          while (this.pageIterator.hasNext) {
            await fetchNextAndCacheForFilesPageIterator(this.pageIterator);
          }
          files = this.pageIterator.value;
          this._preloadedFiles = [];
        }
        filteredByFileExtension = files.filter(file => {
          for (const e of this.fileExtensions) {
            if (e == this.getFileExtension(file.name)) {
              return file;
            }
          }
        });
      }

      if (filteredByFileExtension && filteredByFileExtension.length >= 0) {
        this.files = filteredByFileExtension;
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
  protected handleItemSelect(item: DriveItem, event): void {
    this.fireCustomEvent('itemClick', item);

    // handle accessibility updates when item clicked
    if (event) {
      const fileList = this.renderRoot.querySelector('.file-list');

      // get index of the focused item
      const nodes = Array.from(fileList.children);
      const li = event.target.closest('li');
      const index = nodes.indexOf(li);
      this._focusedItemIndex = index;

      for (let i = 0; i < fileList.children.length; i++) {
        fileList.children[i].classList.remove('focused');
      }
    }
  }

  /**
   * Handle the click event on button to show next page.
   *
   * @protected
   * @memberof MgtFileList
   */
  protected async renderNextPage() {
    const root = this.renderRoot.querySelector('file-list-wrapper');
    if (root && root.animate) {
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

    // render next page from cache if exists, or else use iterator
    if (this._preloadedFiles.length > 0) {
      this.files = [
        ...this.files,
        ...this._preloadedFiles.splice(0, Math.min(this.pageSize, this._preloadedFiles.length))
      ];
    } else {
      if (this.pageIterator.hasNext) {
        await fetchNextAndCacheForFilesPageIterator(this.pageIterator);
        this.files = this.pageIterator.value;
      }
    }

    this.requestUpdate();
  }

  /**
   * Get file extension string from file name
   *
   * @param name file name
   * @returns {string} file extension
   */
  private getFileExtension(name) {
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
  private updateItemBackgroundColor(fileList, focusedItem, className) {
    // reset background color
    for (let i = 0; i < fileList.children.length; i++) {
      fileList.children[i].classList.remove(className);
    }

    // set focused item background color
    if (focusedItem) {
      focusedItem.classList.add(className);
      focusedItem.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'start' });
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
      clearFilesCache();
    }

    this.requestStateUpdate(true);
  }
}
