/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { MgtTemplatedComponent, Providers, ProviderState } from '@microsoft/mgt-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property, TemplateResult } from 'lit-element';
import { repeat } from 'lit-html/directives/repeat';
import {
  getDriveFilesById,
  getDriveFilesByPath,
  getFiles,
  getFilesByListQuery,
  getFilesByQueries,
  getGroupFilesById,
  getMyInsightsFiles,
  getSiteFilesById,
  getUserFilesById,
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
   * @type {[string[]}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'file-queries'
  })
  public get fileQueries(): string[] {
    return this._fileQueries;
  }
  public set fileQueries(value: string[]) {
    if (value === this._fileQueries) {
      return;
    }

    this._fileQueries = value;
    this.requestStateUpdate();
  }

  /**
   * allows developer to provide an array of files
   *
   * @type {[MicrosoftGraph.DriveItem[]}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'files'
  })
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
   * A number value to indicate the maximum number of files to show
   * @type {number}
   */
  @property({
    attribute: 'show-max',
    type: Number
  })
  public showMax: number;

  private _fileListQuery: string;
  private _fileQueries: string[];
  private _files: MicrosoftGraph.DriveItem[];
  private _siteId: string;
  private _itemId: string;
  private _driveId: string;
  private _itemPath: string;
  private _groupId: string;
  private _insightType: OfficeGraphInsightString;
  private _userId: string;

  constructor() {
    super();

    this.showMax = 5;
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
   * Render the list of people.
   *
   * @protected
   * @param {*} files
   * @returns {TemplateResult}
   * @memberof mgtFileList
   */
  protected renderFiles(): TemplateResult {
    console.log('files', this.files);
    const maxFiles = this.files.slice(0, this.showMax);

    return html`
      <ul class="file-list">
        ${repeat(
          maxFiles,
          f => f.id,
          f => html`
            <li class="file-item">
              ${this.renderFile(f)}
            </li>
          `
        )}
        ${this.files.length > this.showMax ? this.renderOverflow() : null}
      </ul>
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
   * Render the overflow content to represent any extra files, beyond the max.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtFileList
   */
  protected renderOverflow(): TemplateResult {
    const extra = this.files.length - this.showMax;
    return (
      this.renderTemplate('overflow', {
        extra,
        max: this.showMax,
        files: this.files
      }) ||
      html`
        <li class="overflow"><span>+${extra}<span></li>
      `
    );
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

    if (
      (this.driveId && (!this.itemId && !this.itemPath)) ||
      (this.groupId && !this.itemId) ||
      (this.siteId && !this.itemId) ||
      (this.userId && (!this.insightType && !this.itemId))
    ) {
      this.files = null;
    }

    const graph = provider.graph.forComponent(this);
    let files;

    if (!this.files) {
      if (this.fileListQuery) {
        files = await getFilesByListQuery(graph, this.fileListQuery);
      } else if (this.fileQueries) {
        files = await getFilesByQueries(graph, this.fileQueries);
      } else if (this.driveId) {
        if (this.itemId) {
          files = await getDriveFilesById(graph, this.driveId, this.itemId);
        } else if (this.itemPath) {
          files = await getDriveFilesByPath(graph, this.driveId, this.itemPath);
        }
      } else if (this.groupId) {
        files = await getGroupFilesById(graph, this.groupId, this.itemId);
      } else if (this.siteId) {
        files = await getSiteFilesById(graph, this.siteId, this.itemId);
      } else if (this.userId) {
        if (this.itemId) {
          files = await getUserFilesById(graph, this.userId, this.itemId);
        } else if (this.insightType) {
          files = await getUserInsightsFiles(graph, this.userId, this.insightType);
        }
      } else if (this.insightType && !this.userId) {
        files = await getMyInsightsFiles(graph, this.insightType);
      } else {
        files = await getFiles(graph);
      }
      this.files = files;
    }
  }
}
