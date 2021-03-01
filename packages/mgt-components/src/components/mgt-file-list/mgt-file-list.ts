/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { arraysAreEqual, MgtTemplatedComponent, Providers, ProviderState } from '@microsoft/mgt-element';
import { DriveItem } from '@microsoft/microsoft-graph-types';
import { customElement, html, internalProperty, property, TemplateResult } from 'lit-element';
import { repeat } from 'lit-html/directives/repeat';
import { classMap } from 'lit-html/directives/class-map';
import {
  getDriveFilesById,
  getDriveFilesByPath,
  getFiles,
  getFilesById,
  getFilesByListQuery,
  getFilesByPath,
  getFilesByQueries,
  getGroupFilesById,
  getGroupFilesByPath,
  getMyInsightsFiles,
  getSiteFilesById,
  getSiteFilesByPath,
  getUserFilesById,
  getUserFilesByPath,
  getUserInsightsFiles
} from '../../graph/graph.files';

import { OfficeGraphInsightString, ViewType } from '../../graph/types';
import { styles } from './mgt-file-list-css';
import { debounce } from '../../utils/Utils';

/**
 * The File List component displays a list of multiple folders and files by using the file/folder name, an icon, and other properties specicified by the developer. This component uses the mgt-file component.
 *
 * @export
 * @class MgtFileList
 * @extends {MgtTemplatedComponent}
 *
 * @fires fileSelected
 *
 * @cssprop
 *
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
    this.requestStateUpdate();
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
      return value.split(',').map(v => v.trim());
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
    this.requestStateUpdate();
  }

  /**
   * allows developer to provide an array of files
   *
   * @type {MicrosoftGraph.DriveItem[]}
   * @memberof MgtFileList
   */
  @property({ type: Object })
  public get files(): DriveItem[] {
    return this._files;
  }
  public set files(value: DriveItem[]) {
    if (value === this._files) {
      return;
    }

    this._files = value;
    this.requestStateUpdate();
  }

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
    this.requestStateUpdate();
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
    this.requestStateUpdate();
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
    this.requestStateUpdate();
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
    this.requestStateUpdate();
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
    this.requestStateUpdate();
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
    this.requestStateUpdate();
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
    this.requestStateUpdate();
  }

  /**
   * A number value to indicate the maximum number of files to show
   * @type {number}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'show-max',
    type: Number
  })
  public showMax: number;

  /**
   * Gets or sets whether expanded section of files are rendered
   *
   * @type {boolean}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'is-expanded',
    type: Boolean
  })
  public isExpanded: boolean;

  /**
   * A boolean value to indicate whether to render more files when scrolled to the bottom
   * @type {number}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'render-on-scroll',
    type: Boolean
  })
  public renderOnScroll: boolean;

  /**
   * The selected item
   *
   * @readonly
   * @memberof MgtFileList
   * @type {MicrosoftGraph.DriveItem}
   */
  public get selectedItem() {
    return this._selectedItem;
  }

  @internalProperty()
  private _selectedItem: DriveItem;

  private _fileListQuery: string;
  private _fileQueries: string[];
  private _files: DriveItem[];
  private _siteId: string;
  private _itemId: string;
  private _driveId: string;
  private _itemPath: string;
  private _groupId: string;
  private _insightType: OfficeGraphInsightString;
  private _userId: string;

  constructor() {
    super();

    this.showMax = 10;
    this._selectedItem = null;
    this.renderOnScroll = false;
  }

  public render() {
    if (!this.files && this.isLoadingState) {
      return this.renderLoading();
    }

    if (!this.files || this.files.length === 0) {
      return this.renderNoData();
    }

    return this.renderTemplate('default', { files: this.files, max: this.showMax }) || this.renderFiles();
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
    const maxFiles = this.files.slice(0, this.showMax);

    const expandedFilesTemplate = this.isExpanded ? this.renderOverflowContent() : this.renderOverflowButton();

    const fileListClasses = classMap({
      'file-list': true,
      scrollable: this.renderOnScroll
    });

    return html`
      <div id="file-list-wrapper" class="file-list-wrapper" onscroll="${this.handleScroll()}">
        <ul id="file-list" class=${fileListClasses}>
          ${repeat(
            maxFiles,
            f => f.id,
            f =>
              html`
                <li class="file-item">${this.renderFile(f)}</li>
              `
          )}
          ${this.files.length > this.showMax ? expandedFilesTemplate : null}
        </ul>

        <div id="file-list-overflow" class="file-list-overflow"></div>
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
    const view = ViewType.twolines;

    const fileClasses = classMap({
      'file-item--selected': this._selectedItem && file.id === this._selectedItem.id
    });

    return (
      this.renderTemplate('file', { file }) ||
      html`
        <mgt-file class=${fileClasses} .fileDetails=${file} .view=${view}></mgt-file>
      `
    );
  }

  /**
   * Render the overflow content to represent any extra files, beyond the max.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtFileList
   */
  protected renderOverflowContent(): TemplateResult {
    if (!this.files && this.isLoadingState) {
      return html`
        <div class="loading">
          <mgt-spinner></mgt-spinner>
        </div>
      `;
    }

    const remainingFiles = this.files.slice(this.showMax);

    return html`
          ${repeat(
            remainingFiles,
            f => f.id,
            f =>
              html`
                <li class="file-item">${this.renderFile(f)}</li>
              `
          )}
      </div>
    `;
  }

  /**
   * Render the overflow button.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtFileList
   */
  protected renderOverflowButton(): TemplateResult {
    const extra = this.files.length - this.showMax;

    if (this.renderOnScroll) {
      return (
        this.renderTemplate('overflow', {
          extra,
          max: this.showMax,
          files: this.files
        }) ||
        html`
          <li class="show-more" @click=${() => this.showRemainingFiles()}><span>Show ${extra} more items<span></li>
        `
      );
    } else {
      return html`
      <li class="show-more"><span>${extra} more items<span></li>
    `;
    }
  }

  /**
   * Handle the click event on an item.
   *
   * @protected
   * @param {MicrosoftGraph.DriveItem} item
   * @memberof MgtFileList
   */
  protected handleItemSelect(item: DriveItem, event: PointerEvent): void {
    if (this._selectedItem === item) {
      this._selectedItem = null;
    } else {
      this._selectedItem = item;
    }

    this.fireCustomEvent('fileSelected', this._selectedItem);
  }

  /**
   * Display the remaining files.
   *
   * @protected
   * @memberof MgtFileList
   */
  protected showRemainingFiles() {
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
    this.isExpanded = true;

    this.fireCustomEvent('expanded', null, true);
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
    let files;

    const getFromMyDrive = !this.driveId && !this.siteId && !this.groupId && !this.userId;

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
        files = await getFilesByListQuery(graph, this.fileListQuery);
      } else if (this.fileQueries) {
        files = await getFilesByQueries(graph, this.fileQueries);
      } else if (getFromMyDrive) {
        if (this.itemId) {
          files = await getFilesById(graph, this.itemId);
        } else if (this.itemPath) {
          files = await getFilesByPath(graph, this.itemPath);
        } else if (this.insightType) {
          files = await getMyInsightsFiles(graph, this.insightType);
        } else {
          files = await getFiles(graph);
        }
      } else if (this.driveId) {
        if (this.itemId) {
          files = await getDriveFilesById(graph, this.driveId, this.itemId);
        } else if (this.itemPath) {
          files = await getDriveFilesByPath(graph, this.driveId, this.itemPath);
        }
      } else if (this.groupId) {
        if (this.itemId) {
          files = await getGroupFilesById(graph, this.groupId, this.itemId);
        } else if (this.itemPath) {
          files = await getGroupFilesByPath(graph, this.groupId, this.itemPath);
        }
      } else if (this.siteId) {
        if (this.itemId) {
          files = await getSiteFilesById(graph, this.siteId, this.itemId);
        } else if (this.itemPath) {
          files = await getSiteFilesByPath(graph, this.siteId, this.itemPath);
        }
      } else if (this.userId) {
        if (this.itemId) {
          files = await getUserFilesById(graph, this.userId, this.itemId);
        } else if (this.itemPath) {
          files = await getUserFilesByPath(graph, this.userId, this.itemPath);
        } else if (this.insightType) {
          files = await getUserInsightsFiles(graph, this.userId, this.insightType);
        }
      }
      this.files = files;
    }

    // Reset the selected item if it doesn't match any of the new results.
    if (this._selectedItem && this.files.findIndex(v => v.id === this._selectedItem.id) === -1) {
      this._selectedItem = null;
    }
  }

  /**
   * Handle load more files on scrolling at the bottom of the list
   */
  private handleScroll() {
    const scrollable = this.shadowRoot.getElementById('file-list');
    const overflow = this.shadowRoot.getElementById('file-list-overflow');

    if (scrollable) {
      scrollable.addEventListener(
        'scroll',
        debounce(() => {
          scrollable.scrollTop === scrollable.scrollHeight - scrollable.offsetHeight
            ? overflow.classList.add('fadeout')
            : overflow.classList.remove('fadeout');
        }, 10)
      );
    }
  }
}
