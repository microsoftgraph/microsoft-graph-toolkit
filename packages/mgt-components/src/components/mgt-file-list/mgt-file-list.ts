/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  GraphPageIterator,
  Providers,
  ProviderState,
  mgtHtml,
  MgtTemplatedTaskComponent,
  registerComponent,
  customElementHelper
} from '@microsoft/mgt-element';
import { DriveItem, SharedInsight } from '@microsoft/microsoft-graph-types';
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
import { OfficeGraphInsightString, ViewType, viewTypeConverter } from '../../graph/types';
import { styles } from './mgt-file-list-css';
import { strings } from './strings';
import { MgtFile, registerMgtFileComponent } from '../mgt-file/mgt-file';
import { MgtFileUploadConfig, registerMgtFileUploadComponent } from './mgt-file-upload/mgt-file-upload';

import { fluentProgressRing } from '@fluentui/web-components';
import { registerFluentComponents } from '../../utils/FluentComponents';
import { CardSection } from '../BasePersonCardSection';
import { getRelativeDisplayDate } from '../../utils/Utils';
import { getFileTypeIconUri } from '../../styles/fluent-icons';

export const registerMgtFileListComponent = () => {
  registerFluentComponents(fluentProgressRing);

  registerMgtFileComponent();
  registerMgtFileUploadComponent();
  registerComponent('file-list', MgtFileList);
};

const isSharedInsight = (sharedInsightFile: SharedInsight): sharedInsightFile is SharedInsight => {
  return 'lastShared' in sharedInsightFile;
};

/**
 * The File List component displays a list of multiple folders and files by
 * using the file/folder name, an icon, and other properties specified by the developer.
 * This component uses the mgt-file component.
 *
 * @export
 * @class MgtFileList
 *
 * @fires {CustomEvent<undefined>} updated - Fired when the component is updated
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
export class MgtFileList extends MgtTemplatedTaskComponent implements CardSection {
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

  // files from the person card component
  @state()
  private _personCardFiles: DriveItem[];

  /**
   * allows developer to provide query for a file list
   *
   * @type {string}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'file-list-query'
  })
  public fileListQuery: string;

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
  public fileQueries: string[] | null = null;

  /**
   * allows developer to provide an array of files
   *
   * @type {MicrosoftGraph.DriveItem[]}
   * @memberof MgtFileList
   */
  @property({ type: Object })
  public files: DriveItem[] | null = null;

  /**
   * allows developer to provide site id for a file
   *
   * @type {string}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'site-id'
  })
  public siteId: string;

  /**
   * allows developer to provide drive id for a file
   *
   * @type {string}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'drive-id'
  })
  public driveId: string;

  /**
   * allows developer to provide group id for a file
   *
   * @type {string}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'group-id'
  })
  public groupId: string;

  /**
   * allows developer to provide item id for a file
   *
   * @type {string}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'item-id'
  })
  public itemId: string;

  /**
   * allows developer to provide item path for a file
   *
   * @type {string}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'item-path'
  })
  public itemPath: string;

  /**
   * allows developer to provide user id for a file
   *
   * @type {string}
   * @memberof MgtFile
   */
  @property({
    attribute: 'user-id'
  })
  public userId: string;

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
  public insightType: OfficeGraphInsightString;

  /**
   * Sets what data to be rendered (file icon only, oneline, twolines threelines).
   * Default is 'threelines'.
   *
   * @type {ViewType}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'item-view',
    converter: value => viewTypeConverter(value, 'threelines')
  })
  public itemView: ViewType = 'threelines';

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
  public fileExtensions: string[] = [];

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
  public pageSize = 10;

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
  public hideMoreFilesButton = false;

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
  public maxFileSize: number;

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
  public enableFileUpload = false;

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
  public maxUploadFile = 10;

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
  public excludedFileExtensions: string[] = [];

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

  private _preloadedFiles: DriveItem[] = [];
  private pageIterator: GraphPageIterator<DriveItem>;
  // tracking user arrow key input of selection for accessibility purpose
  private _focusedItemIndex = -1;

  @state() private _isLoadingMore: boolean;

  constructor(files?: DriveItem[]) {
    super();
    this._personCardFiles = files;
  }

  /**
   * Reset state
   *
   * @memberof MgtFileList
   */
  protected clearState(): void {
    super.clearState();
    this.files = null;
    this._personCardFiles = null;
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

  protected args(): unknown[] {
    return [
      this.providerState,
      this.fileListQuery,
      this.fileQueries,
      this.siteId,
      this.driveId,
      this.groupId,
      this.itemId,
      this.itemPath,
      this.userId,
      this.insightType,
      this.fileExtensions,
      this.pageSize,
      this.maxFileSize
    ];
  }

  protected renderLoading = () => {
    if (!this.files) {
      return this.renderTemplate('loading', null) || html``;
    }
    return this.renderContent();
  };

  /**
   * Render the file list
   *
   * @return {*}
   * @memberof MgtFileList
   */
  protected renderContent = () => {
    if (!this.files || this.files.length === 0) {
      return this.renderNoData();
    }
    if (this._personCardFiles) {
      this.files = this._personCardFiles;
    }
    return this._isCompact ? this.renderCompactView() : this.renderFullView();
  };

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
            id="file-list-item-${this.files[0].id}"
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
                id="file-list-item-${f.id}"
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
  protected renderFile(file: DriveItem | SharedInsight): TemplateResult {
    const view = this.itemView;
    // if file is type SharedInsight, render Shared Insight File
    if (isSharedInsight(file)) {
      return this.renderSharedInsightFile(file);
    }
    return (
      this.renderTemplate('file', { file }, file.id) ||
      mgtHtml`
        <mgt-file class="mgt-file-item" .fileDetails=${file} .view=${view}></mgt-file>
      `
    );
  }

  /**
   * Render a file item of Shared Insight Type
   *
   * @protected
   * @param {IFile} file
   * @returns {TemplateResult}
   * @memberof MgtFileList
   */
  protected renderSharedInsightFile(file: SharedInsight): TemplateResult {
    const lastModifiedTemplate = file.lastShared
      ? html`
          <div class="shared_insight_file__last-modified">
            ${this.strings.sharedTextSubtitle} ${getRelativeDisplayDate(new Date(file.lastShared.sharedDateTime))}
          </div>
        `
      : null;

    return html`
      <div class="shared_insight_file" @click=${(e: MouseEvent) => this.handleSharedInsightClick(file, e)} tabindex="0">
        <div class="shared_insight_file__icon">
          <img alt="${file.resourceVisualization.title}" src=${getFileTypeIconUri(
            file.resourceVisualization.type,
            48,
            'svg'
          )} />
        </div>
        <div class="shared_insight_file__details">
          <div class="shared_insight_file__name">${file.resourceVisualization.title}</div>
          ${lastModifiedTemplate}
        </div>
      </div>
    `;
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
    const target = event.target as HTMLElement;
    let fileList: HTMLElement;
    if (!target.classList) {
      fileList = this.renderRoot.querySelector('.file-list-children');
    } else {
      fileList = this.renderRoot.querySelector('.file-list');
    }
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
      if (this.fileExtensions?.length > 0) {
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
    for (const file of this.files) {
      if (file?.folder?.childCount > 0) {
        // expand the file with children
        const driveId = file?.parentReference?.driveId;
        const itemId = file?.id;
        const iterator = await getDriveFilesByIdIterator(graph, driveId, itemId, 5);
        if (iterator) {
          const children = [...iterator.value];
          file.children = children;
        }
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

  private readonly handleSharedInsightClick = (file: SharedInsight, e?: MouseEvent) => {
    if (file.resourceReference?.webUrl && !this.disableOpenOnClick) {
      e.preventDefault();
      window.open(file.resourceReference.webUrl, '_blank', 'noreferrer');
    }
  };

  private readonly handleFileClick = (file: DriveItem) => {
    const hasChildFolders = file?.folder?.childCount > 0 && file?.children;
    // the item has child folders, on click should get the child folders and render them
    if (hasChildFolders) {
      this.showChildren(file.id);
      return;
    }

    if (file?.webUrl && !this.disableOpenOnClick) {
      window.open(file.webUrl, '_blank', 'noreferrer');
    }
  };

  private readonly showChildren = (fileId: string) => {
    const itemDOM = this.renderRoot.querySelector(`#file-list-item-${fileId}`);
    this.renderChildren(fileId, itemDOM);
  };

  private readonly renderChildren = (itemId: string, itemDOM: Element) => {
    const fileListName = customElementHelper.isDisambiguated
      ? `${customElementHelper.prefix}-file-list`
      : 'mgt-file-list';
    const childrenContainer = this.renderRoot.querySelector(`#file-list-children-${itemId}`);
    if (!childrenContainer) {
      const fl = document.createElement(fileListName);
      fl.setAttribute('item-id', itemId);
      fl.setAttribute('id', `file-list-children-${itemId}`);
      fl.setAttribute('class', 'file-list-children-show');
      itemDOM.after(fl);
    } else {
      // toggle to show/hide the children container
      if (childrenContainer.classList.contains('file-list-children-hide')) {
        childrenContainer.setAttribute('class', 'file-list-children-show');
      } else {
        childrenContainer.setAttribute('class', 'file-list-children-hide');
      }
    }
  };

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
    // explicitly run the task to reload data
    void this._task.run();
  }
}
