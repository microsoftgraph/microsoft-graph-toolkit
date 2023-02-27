/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { GraphPageIterator, Providers, ProviderState, customElement, mgtHtml } from '@microsoft/mgt-element';
import { DriveItem } from '@microsoft/microsoft-graph-types';
import { html, nothing, TemplateResult } from 'lit';
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
import '../mgt-file-list/mgt-file-upload/mgt-file-upload';
import { ViewType } from '../../graph/types';
import { styles } from './mgt-file-grid-css';
import { strings } from '../mgt-file-list/strings';
import { MgtFile } from '../mgt-file/mgt-file';
import { MgtFileUploadConfig } from '../mgt-file-list/mgt-file-upload/mgt-file-upload';

import { fluentProgressRing, fluentDesignSystemProvider, fluentButton } from '@fluentui/web-components';
import { registerFluentComponents } from '../../utils/FluentComponents';
import { classMap } from 'lit/directives/class-map.js';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';
import '../mgt-menu/mgt-menu';
import { MgtFileListBase } from '../mgt-file-list/mgt-file-list-base';
import { Command } from '../mgt-menu/mgt-menu';
import '../mgt-person/mgt-person';
import { formatBytes, getRelativeDisplayDate } from '../../utils/Utils';

registerFluentComponents(fluentProgressRing, fluentDesignSystemProvider, fluentButton);

// re-export to ensure it's in the final package as mgt-menu is internal only
export { Command };

/**
 * The File List component displays a list of multiple folders and files by
 * using the file/folder name, an icon, and other properties specified by the developer.
 * This component uses the mgt-file component.
 *
 * @export
 * @class MgtFileList
 * @extends {MgtTemplatedComponent}
 *
 * @fires {CustomEvent<MicrosoftGraph.DriveItem>} itemClick - Fired when user click a file. Returns the file (DriveItem) details.
 * @fires {CustomEvent<MicrosoftGraph.DriveItem[]>} selectionChanged - Fired when user select a file. Returns the selected files (DriveItem) details.
 *
 * @cssprop --file-upload-border- {String} File upload border top style
 * @cssprop --file-upload-background-color - {Color} File upload background color with opacity style
 * @cssprop --file-upload-button-float - {string} Upload button float position
 * @cssprop --file-upload-button-background-color - {Color} Background color of upload button
 * @cssprop --file-upload-dialog-background-color - {Color} Background color of upload dialog
 * @cssprop --file-upload-dialog-content-background-color - {Color} Background color of dialog content
 * @cssprop --file-upload-dialog-content-color - {Color} Color of dialog content
 * @cssprop --file-upload-dialog-primarybutton-background-color - {Color} Background color of primary button
 * @cssprop --file-upload-dialog-primarybutton-color - {Color} Color text of primary button
 * @cssprop --file-upload-button-color - {Color} Text color of upload button
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
 * @cssprop --progress-ring-size -{String} Progress ring height and width
 */

// tslint:disable-next-line: max-classes-per-file
@customElement('file-grid')
export class MgtFileGrid extends MgtFileListBase {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * Strings to be used in the component
   *
   * @readonly
   * @protected
   * @memberof MgtFileList
   */
  protected get strings() {
    return strings;
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

  /**
   * Property to set the available actions on the file context menu
   *
   * @type {Command<MicrosoftGraph.DriveItem>[]}
   * @memberof MgtFileGrid
   */
  @property({
    attribute: false
  })
  public commands: Command<DriveItem>[] = [];

  private _preloadedFiles: DriveItem[];
  private pageIterator: GraphPageIterator<DriveItem>;
  // tracking user arrow key input of selection for accessibility purpose
  private _focusedItemIndex: number = -1;

  @state()
  private _isLoadingMore: boolean;

  @state()
  private _selectedFiles: Map<string, DriveItem>;

  constructor() {
    super();
    this._selectedFiles = new Map();
    this.pageSize = 10;
    this.itemView = ViewType.image;
    this.maxUploadFile = 10;
    this.enableFileUpload = false;
    this._preloadedFiles = [];
  }

  /**
   * Override requestStateUpdate to include clearstate.
   *
   * @memberof MgtFileGrid
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
    this._selectedFiles = new Map();
    this.fireCustomEvent('selectionChanged', []);
  }

  /**
   * Render the file list
   *
   * @memberof MgtFileList
   */
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
   * @memberof MgtFileList
   */
  protected renderLoading() {
    return this.renderTemplate('loading', null) || nothing;
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
          <fluent-design-system-provider use-defaults>
            <div id="file-list-wrapper" class="file-list-wrapper" dir=${this.direction}>
              ${this.renderFileUpload()}
            </div>
          </fluent-design-system-provider>
      `
        : html``)
    );
  }

  /**
   * render the files in a data grid
   *
   * @protected
   * @return {*}  {TemplateResult}
   * @memberof MgtFileList
   */
  protected renderFiles(): TemplateResult {
    const headerClasses = {
      header: true,
      selected: this.allSelected()
    };
    // the hidden anchor tag is used to download file
    return html`
      <a id="file-link" style="display:none"  target="_blank"></a>
      <div id="file-list-wrapper" class="file-grid"  dir=${this.direction}>
        <div class="file-row">
          <div class=${classMap(headerClasses)}>${this.renderSelectorHeader()}</div>
          <div class=${classMap(headerClasses)}></div>
          <div class=${classMap(headerClasses)}>Name</div>
          <div class=${classMap(headerClasses)}>Modified</div>
          <div class=${classMap(headerClasses)}>Modified By</div>
          <div class=${classMap(headerClasses)}>Size</div>
        </div>
        ${repeat(
          this.files,
          f => f.id,
          f => html`
                <div
                  class="${this.isSelected(f) ? 'file-row selected' : 'file-row'}"
                  @click=${() => this.onSelectorClicked(f)}
                >
                  <div class="cell">${this.renderSelector(f)}</div>
                  <div class="cell">${this.renderFileIcon(f)}</div>
                  <div class="cell file-name">
                    ${this.renderFileName(f)}
                    ${this.renderMenu(f)}
                  </div>
                  <div class="cell">${getRelativeDisplayDate(new Date(f.lastModifiedDateTime))}</div>
                  <div class="cell">${this.renderUser(f)}</div>
                  <div class="cell">${this.sizeText(f)}</div>
                </div>
              `
        )}
      </div>
        ${
          !this.hideMoreFilesButton && this.pageIterator && (this.pageIterator.hasNext || this._preloadedFiles.length)
            ? this.renderMoreFileButton()
            : null
        }
    `;
  }

  private sizeText(file: DriveItem) {
    if (file.folder) return `${file.folder.childCount} ${file.folder.childCount === 0 ? strings.item : strings.items}`;

    return formatBytes(file.size);
  }

  private renderUser(file: DriveItem) {
    if (file.lastModifiedByUser)
      return mgtHtml`
      <mgt-person
        show-presence
        avatar-size="small"
        view="oneline"
        .personDetails=${file.lastModifiedByUser}
      ></mgt-person>`;

    if (file.lastModifiedBy?.user)
      return mgtHtml`
      <mgt-person
        show-presence
        avatar-size="small"
        view="oneline"
        user-id=${file.lastModifiedBy.user.id}
      ></mgt-person>`;

    return nothing;
  }

  private isSelected(file: DriveItem): boolean {
    return this._selectedFiles.has(file.id);
  }

  private renderSelectorHeader(): TemplateResult {
    const classes = {
      'file-selector': true,
      selected: this.allSelected()
    };

    return html`
      <div class=${classMap(classes)} @click=${this.allSelected() ? this.deselectAll : this.selectAll}>
        ${getSvg(SvgIcon.FilledCheckMark)}
      </div>
    `;
  }

  private allSelected(): boolean {
    return this.files && this.files.length > 0 && this._selectedFiles.size === this.files.length;
  }

  private selectAll(): void {
    const tmp = new Map();
    this.files.forEach(file => tmp.set(file.id, file));
    this._selectedFiles = tmp;
    this.fireCustomEvent('selectionChanged', this.files);
  }

  private deselectAll(): void {
    this._selectedFiles = new Map();
    this.fireCustomEvent('selectionChanged', []);
  }

  private renderSelector(file: DriveItem): TemplateResult {
    const classes = {
      'file-selector': true,
      selected: this.isSelected(file)
    };

    return html`<div class=${classMap(classes)}>${getSvg(SvgIcon.FilledCheckMark)}</div>`;
  }

  private onSelectorClicked(file: DriveItem): void {
    if (this._selectedFiles.has(file.id)) {
      this._selectedFiles.delete(file.id);
    } else {
      this._selectedFiles.set(file.id, file);
    }

    // request a re-render as we're mutating the state of the _selectedFiles map without an assignment
    this.requestUpdate();

    this.fireCustomEvent(
      'selectionChanged',
      Array.from(this._selectedFiles, ([, value]) => value)
    );
  }

  private renderMenu(file: DriveItem): TemplateResult {
    return html`${
      !this.commands || this.commands.length === 0
        ? nothing
        : mgtHtml`
          <mgt-menu .commands=${this.commands} .item=${file}></mgt-menu>
        `
    }`;
  }

  /**
   * Render an individual file.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof mgtFileList
   */
  protected renderFileIcon(file: DriveItem): TemplateResult {
    const view = this.itemView;
    return (
      this.renderTemplate('file', { file }, file.id) ||
      mgtHtml`
        <mgt-file
          @click=${e => this.handleItemSelect(file, e)}
          class="file-item"
          .fileDetails=${file}
          .view=${view}
        ></mgt-file>
      `
    );
  }

  /**
   * Render an individual file.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof mgtFileList
   */
  protected renderFileName(file: DriveItem): TemplateResult {
    return file.folder
      ? html`<a class="file-item" href="#" @click=${e => this.handleItemSelect(file, e)}>${file.name}</a>`
      : html`<a class="file-item" target="_blank" href=${file.webUrl}>${file.name}</a>`;
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
        <div class="loading-indicator">
          <fluent-progress-ring role="progressbar" viewBox="0 0 8 8" class="progress-ring"></fluent-progress-ring>
        </div>
      `;
    } else {
      return html`
        <fluent-button
          appearance="mgt-file-grid.scss"
          id="show-more"
          class="show-more"
          @click=${() => this.renderNextPage()}
        >
          ${this.strings.showMoreSubtitle}
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
    let focusedItem: Element;

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

      const file = focusedItem.children[0] as MgtFile;
      event.preventDefault();
      this.raiseItemClickedEvent(file.fileDetails);

      this.updateItemBackgroundColor(fileList, focusedItem, 'selected');
    }

    if (event.code === 'Tab') {
      focusedItem = fileList.children[this._focusedItemIndex];
      focusedItem?.classList.remove('focused');
    }
  }

  private raiseItemClickedEvent(file: DriveItem) {
    this.fireCustomEvent('itemClick', file);
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
        // retrieve all pages before filtering
        if (this.pageIterator && this.pageIterator.value) {
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

      if (filteredByFileExtension && filteredByFileExtension.length >= 0) {
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
  protected handleItemSelect(item: DriveItem, event: MouseEvent): void {
    event?.stopPropagation();
    this.raiseItemClickedEvent(item);
    if (item.file && item.webUrl) {
      // open the web url if the item is a file
      this.clickFileLink(item.webUrl);
    }
  }

  private clickFileLink = (url: string) => {
    const a = this.renderRoot.querySelector('#file-link') as HTMLAnchorElement;
    a.href = url;
    if (a.href) {
      a.click();
      a.href = '';
    }
  };

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
        await fetchNextAndCacheForFilesPageIterator(this.pageIterator);
        this._isLoadingMore = false;
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
