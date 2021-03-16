/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { arraysAreEqual, MgtTemplatedComponent, Providers, ProviderState } from '@microsoft/mgt-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property, TemplateResult } from 'lit-element';
import { debounce } from '../../utils/Utils';
import { repeat } from 'lit-html/directives/repeat';
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

/**
 * The File List component displays a list of multiple folders and files by using the file/folder name, an icon, and other properties specicified by the developer. This component uses the mgt-file component.
 *
 * @export
 * @class MgtFileList
 * @extends {MgtTemplatedComponent}
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
  public get files(): MicrosoftGraph.DriveItem[] {
    return this._files;
  }
  public set files(value: MicrosoftGraph.DriveItem[]) {
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
   * A number value to indicate the number of more files to load when show more button is clicked
   * @type {number}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'load-more-file-count',
    type: Number
  })
  public loadMoreFileCount: number;

  private _fileListQuery: string;
  private _fileQueries: string[];
  private _files: MicrosoftGraph.DriveItem[];
  private _siteId: string;
  private _itemId: string;
  private _driveId: string;
  private _itemPath: string;
  private _groupId: string;
  private _insightType: OfficeGraphInsightString;
  private _fileExtensions: string[];
  private _userId: string;
  private _renderedFileCount: number;
  private _filesToRender: MicrosoftGraph.DriveItem[];

  constructor() {
    super();

    this.showMax = 10;
    this._renderedFileCount = this.showMax;
    this.loadMoreFileCount = 5;
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
      <div id="file-list-wrapper" class="file-list-wrapper">
        <ul id="file-list" class="file-list">
          ${repeat(
            this._filesToRender,
            f => f.id,
            f => html`
              <li class="file-item">
                ${this.renderFile(f)}
              </li>
            `
          )}
          ${this._renderedFileCount < this.files.length ? this.renderMoreFileButton() : null}
        </ul>
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
  protected renderFile(file: MicrosoftGraph.DriveItem): TemplateResult {
    const view = ViewType.twolines;
    return (
      this.renderTemplate('file', { file }) ||
      html`
        <mgt-file .fileDetails=${file} .view=${view}></mgt-file>
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
    const extra = this.files.length - this._renderedFileCount;
    return html`
        <li class="show-more" @click=${() =>
          this.showMoreFiles(this.loadMoreFileCount)}><span>${extra} more items<span></li>
      `;
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
    let files: MicrosoftGraph.DriveItem[];

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
      if (this.fileListQuery && this.fileListQuery !== null) {
        files = await getFilesByListQuery(graph, this.fileListQuery);
      } else if (this.fileQueries && this.fileQueries !== null) {
        files = await getFilesByQueries(graph, this.fileQueries);
      } else if (getFromMyDrive) {
        if (this.itemId && this.itemId !== null) {
          files = await getFilesById(graph, this.itemId);
        } else if (this.itemPath && this.itemPath !== null) {
          files = await getFilesByPath(graph, this.itemPath);
        } else if (this.insightType && this.insightType !== null) {
          files = await getMyInsightsFiles(graph, this.insightType);
        } else {
          files = await getFiles(graph);
        }
      } else if (this.driveId && this.driveId !== null) {
        if (this.itemId && this.itemId !== null) {
          files = await getDriveFilesById(graph, this.driveId, this.itemId);
        } else if (this.itemPath && this.itemPath !== null) {
          files = await getDriveFilesByPath(graph, this.driveId, this.itemPath);
        }
      } else if (this.groupId && this.groupId !== null) {
        if (this.itemId && this.itemId !== null) {
          files = await getGroupFilesById(graph, this.groupId, this.itemId);
        } else if (this.itemPath && this.itemPath !== null) {
          files = await getGroupFilesByPath(graph, this.groupId, this.itemPath);
        }
      } else if (this.siteId && this.siteId !== null) {
        if (this.itemId && this.itemId !== null) {
          files = await getSiteFilesById(graph, this.siteId, this.itemId);
        } else if (this.itemPath && this.itemPath !== null) {
          files = await getSiteFilesByPath(graph, this.siteId, this.itemPath);
        }
      } else if (this.userId && this.userId !== null) {
        if (this.itemId && this.itemId !== null) {
          files = await getUserFilesById(graph, this.userId, this.itemId);
        } else if (this.itemPath && this.itemPath !== null) {
          files = await getUserFilesByPath(graph, this.userId, this.itemPath);
        } else if (this.insightType && this.insightType !== null) {
          files = await getUserInsightsFiles(graph, this.userId, this.insightType);
        }
      }

      this.files = files;

      // filter files when extensions are provided
      let filteredByFileExtension: MicrosoftGraph.DriveItem[];
      if (this.fileExtensions && this.fileExtensions !== null) {
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
      }

      this._filesToRender = this.files.slice(0, this.showMax);
    }
  }

  private showMoreFiles(count: number) {
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
    if (this.files.length > this._renderedFileCount) {
      if (this.files.length - this._renderedFileCount > count) {
        this._filesToRender = this.files.slice(0, this._renderedFileCount + count);
        this._renderedFileCount += count;
        this.requestUpdate();
      } else {
        this._renderedFileCount = this.files.length;
        this._filesToRender = this.files.slice(0, this.files.length);
        this.requestUpdate();
      }
    }
  }

  private getFileExtension(name) {
    const re = /(?:\.([^.]+))?$/;
    const fileExtension = re.exec(name)[1] || '';

    return fileExtension;
  }
}
