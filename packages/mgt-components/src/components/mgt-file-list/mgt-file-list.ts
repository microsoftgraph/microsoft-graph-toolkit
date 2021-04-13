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
import { Drive, DriveItem } from '@microsoft/microsoft-graph-types';
import { customElement, html, property, TemplateResult, internalProperty } from 'lit-element';
import { repeat } from 'lit-html/directives/repeat';
import { FluentProgressRing } from '@fluentui/web-components/dist/web-components.min';

import {
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
import { OfficeGraphInsightString, ViewType } from '../../graph/types';
import { styles } from './mgt-file-list-css';

console.log('This is a temporary workaround for using @fluentui/web-components', FluentProgressRing.name);

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
  public pageSize: number;

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

  private _fileListQuery: string;
  private _fileQueries: string[];
  private _siteId: string;
  private _itemId: string;
  private _driveId: string;
  private _itemPath: string;
  private _groupId: string;
  private _insightType: OfficeGraphInsightString;
  private _fileExtensions: string[];
  private _userId: string;
  private pageIterator: GraphPageIterator<DriveItem>;

  @internalProperty() private _isLoadingMore: boolean;

  constructor() {
    super();

    this.pageSize = 10;
    this.itemView = ViewType.twolines;
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
        <ul id="file-list" class="file-list">
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
        ${this.renderMoreFileButton()}
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
      this.renderTemplate('file', { file }) ||
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
    if (this._isLoadingMore) {
      return html`
        <fluent-design-system-provider use-defaults>
          <fluent-progress-ring class="show-more-progress"></fluent-progress-ring>
        </fluent-design-system-provider>
      `;
    } else if (!this.hideMoreFilesButton && this.pageIterator && this.pageIterator.hasNext) {
      return html`<a id="show-more" class="show-more" @click=${() =>
        this.renderNextPage()}><span>Show more items<span></a>`;
    }
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
      (this.driveId && (!this.itemId && !this.itemPath)) ||
      (this.groupId && (!this.itemId && !this.itemPath)) ||
      (this.siteId && (!this.itemId && !this.itemPath)) ||
      (this.userId && (!this.insightType && (!this.itemId && !this.itemPath)))
    ) {
      this.files = null;
    }

    if (!this.files) {
      if (this.fileListQuery) {
        pageIterator = await getFilesByListQueryIterator(graph, this.fileListQuery, this.pageSize);
        files = pageIterator ? pageIterator.value : null;
      } else if (this.fileQueries) {
        files = await getFilesByQueries(graph, this.fileQueries);
      } else if (getFromMyDrive) {
        if (this.itemId) {
          pageIterator = await getFilesByIdIterator(graph, this.itemId, this.pageSize);
          files = pageIterator ? pageIterator.value : null;
        } else if (this.itemPath) {
          pageIterator = await getFilesByPathIterator(graph, this.itemPath, this.pageSize);
          files = pageIterator ? pageIterator.value : null;
        } else if (this.insightType) {
          files = await getMyInsightsFiles(graph, this.insightType);
        } else {
          pageIterator = await getFilesIterator(graph, this.pageSize);
          files = pageIterator ? pageIterator.value : null;
        }
      } else if (this.driveId) {
        if (this.itemId) {
          pageIterator = await getDriveFilesByIdIterator(graph, this.driveId, this.itemId, this.pageSize);
          files = pageIterator ? pageIterator.value : null;
        } else if (this.itemPath) {
          pageIterator = await getDriveFilesByPathIterator(graph, this.driveId, this.itemPath, this.pageSize);
          files = pageIterator ? pageIterator.value : null;
        }
      } else if (this.groupId) {
        if (this.itemId) {
          pageIterator = await getGroupFilesByIdIterator(graph, this.groupId, this.itemId, this.pageSize);
          files = pageIterator ? pageIterator.value : null;
        } else if (this.itemPath) {
          pageIterator = await getGroupFilesByPathIterator(graph, this.groupId, this.itemPath, this.pageSize);
          files = pageIterator ? pageIterator.value : null;
        }
      } else if (this.siteId) {
        if (this.itemId) {
          pageIterator = await getSiteFilesByIdIterator(graph, this.siteId, this.itemId, this.pageSize);
          files = pageIterator ? pageIterator.value : null;
        } else if (this.itemPath) {
          pageIterator = await getSiteFilesByPathIterator(graph, this.siteId, this.itemPath, this.pageSize);
          files = pageIterator ? pageIterator.value : null;
        }
      } else if (this.userId) {
        if (this.itemId) {
          pageIterator = await getUserFilesByIdIterator(graph, this.userId, this.itemId, this.pageSize);
          files = pageIterator ? pageIterator.value : null;
        } else if (this.itemPath) {
          pageIterator = await getUserFilesByPathIterator(graph, this.userId, this.itemPath, this.pageSize);
          files = pageIterator ? pageIterator.value : null;
        } else if (this.insightType) {
          files = await getUserInsightsFiles(graph, this.userId, this.insightType);
        }
      }

      if (pageIterator) {
        this.pageIterator = pageIterator;
      }

      // filter files when extensions are provided
      let filteredByFileExtension: DriveItem[];
      if (this.fileExtensions && this.fileExtensions !== null) {
        // retrive all pages before filtering
        if (pageIterator && pageIterator.value) {
          while (pageIterator.hasNext) {
            await pageIterator.next();
          }
          files = pageIterator.value;
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
  protected handleItemSelect(item: DriveItem, event: PointerEvent): void {
    this.fireCustomEvent('itemClick', item);
  }

  /**
   * Handle the click event on button to show next page.
   *
   * @protected
   * @memberof MgtFileList
   */
  protected async renderNextPage() {
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

      await this.pageIterator.next();
      this._isLoadingMore = false;
      this.files = this.pageIterator.value;
    }
  }

  private getFileExtension(name) {
    const re = /(?:\.([^.]+))?$/;
    const fileExtension = re.exec(name)[1] || '';

    return fileExtension;
  }
}
