import { customElement, mgtHtml } from '@microsoft/mgt-element';
import { DriveItem } from '@microsoft/microsoft-graph-types';
import { nothing } from 'lit';
import { property, state } from 'lit/decorators.js';
import { OfficeGraphInsightString } from '../../graph/types';
import { BreadcrumbInfo } from '../mgt-breadcrumb/mgt-breadcrumb';
import { MgtFileListBase } from '../mgt-file-list/mgt-file-list-base';
import { strings } from './strings';
import '../mgt-file-grid/mgt-file-grid';

/**
 * FileListBreadCrumb interface
 */
type FileListBreadCrumb = {
  // tslint:disable: completed-docs
  name: string;
  fileListQuery?: string;
  itemId?: string;
  itemPath?: string;
  files?: DriveItem[];
  fileQueries?: string[];
  groupId?: string;
  driveId?: string;
  siteId?: string;
  userId?: string;
  insightType?: OfficeGraphInsightString;
  fileExtensions?: string[];
  // tslint:enable: completed-docs
} & BreadcrumbInfo;

/**
 * A File list composite component
 * Provides a breadcrumb navigation to support folder navigation
 *
 * @class MgtFileComposite
 * @extends {MgtTemplatedComponent}
 */
@customElement('file-list-composite')
class MgtFileListComposite extends MgtFileListBase {
  constructor() {
    super();
    this.breadcrumbRootName = strings.rootNode;
  }

  /**
   * Name to be used for the root node of the breadcrumb
   *
   * @type {string}
   * @memberof MgtFileComposite
   */
  @property({
    attribute: 'breadcrumb-root-name'
  })
  public breadcrumbRootName: string;

  /**
   * Switch to use grid view instead of list view
   *
   * @type {boolean}
   * @memberof MgtFileListComposite
   */
  @property({
    attribute: 'use-grid-view',
    type: Boolean
  })
  public useGridView: boolean;

  /**
   * Override connectedCallback to set initial breadcrumbstate.
   *
   * @memberof MgtFileList
   */
  public connectedCallback(): void {
    super.connectedCallback();
    this.breadcrumb.push({
      name: this.breadcrumbRootName,
      siteId: this.siteId,
      groupId: this.groupId,
      driveId: this.driveId,
      userId: this.userId,
      files: this.files,
      fileExtensions: this.fileExtensions,
      fileListQuery: this.fileListQuery,
      fileQueries: this.fileQueries,
      itemPath: this.itemPath,
      insightType: this.insightType,
      itemId: this.itemId,
      id: 'root-item'
    });
  }

  /**
   * Strings for localization
   *
   * @readonly
   * @protected
   * @memberof MgtFileComposite
   */
  protected get strings() {
    return strings;
  }

  /**
   * An array of nodes to show in the breadcrumb
   *
   * @type {BreadcrumbInfo[]}
   * @readonly
   * @memberof MgtFileList
   */
  @state()
  private breadcrumb: FileListBreadCrumb[] = [];

  /**
   * Render the component
   *
   * @return {*}
   * @memberof MgtFileComposite
   */
  public render() {
    return mgtHtml`
      ${this.renderBreadcrumb()}
      ${this.renderFiles()}
    `;
  }

  private renderFiles() {
    return this.useGridView
      ? mgtHtml`
        <mgt-file-grid
          .fileListQuery=${this.fileListQuery || nothing}
          .itemId=${this.itemId || nothing}
          .itemPath=${this.itemPath || nothing}
          .files=${this.files || nothing}
          .fileQueries=${this.files || nothing}
          .groupId=${this.groupId || nothing}
          .driveId=${this.driveId || nothing}
          .siteId=${this.siteId || nothing}
          .userId=${this.userId || nothing}
          .insightType=${this.insightType || nothing}
          .fileExtensions=${this.fileExtensions || nothing}
          .itemView=${this.itemView || nothing}
          .pageSize=${this.pageSize || nothing}
          .hideMoreFilesButton=${this.hideMoreFilesButton || nothing}
          .maxFileSize=${this.maxFileSize || nothing}
          .enableFileUpload=${this.enableFileUpload || nothing}
          .maxUploadFile=${this.maxUploadFile || nothing}
          .excludedFileExtensions=${this.excludedFileExtensions || nothing}
          @itemClick=${this.handleItemClick}
        ></mgt-file-grid>
`
      : mgtHtml`
        <mgt-file-list
          .fileListQuery=${this.fileListQuery || nothing}
          .itemId=${this.itemId || nothing}
          .itemPath=${this.itemPath || nothing}
          .files=${this.files || nothing}
          .fileQueries=${this.files || nothing}
          .groupId=${this.groupId || nothing}
          .driveId=${this.driveId || nothing}
          .siteId=${this.siteId || nothing}
          .userId=${this.userId || nothing}
          .insightType=${this.insightType || nothing}
          .fileExtensions=${this.fileExtensions || nothing}
          .itemView=${this.itemView || nothing}
          .pageSize=${this.pageSize || nothing}
          .hideMoreFilesButton=${this.hideMoreFilesButton || nothing}
          .maxFileSize=${this.maxFileSize || nothing}
          .enableFileUpload=${this.enableFileUpload || nothing}
          .maxUploadFile=${this.maxUploadFile || nothing}
          .excludedFileExtensions=${this.excludedFileExtensions || nothing}
          @itemClick=${this.handleItemClick}
        ></mgt-file-list>
`;
  }

  private renderBreadcrumb() {
    return mgtHtml`
      <mgt-breadcrumb
        .breadcrumb=${this.breadcrumb.slice()}
        @breadcrumbclick=${this.handleBreadcrumbClick}
      ></mgt-breadcrumb>
`;
  }

  private handleItemClick(e: CustomEvent<DriveItem>): void {
    const item = e.detail;
    if (item.folder) {
      // load folder contents, update breadcrumb
      this.breadcrumb = [...this.breadcrumb, { name: item.name, itemId: item.id, id: item.id }];
      // clear any existing query properties
      this.siteId = null;
      this.groupId = null;
      this.driveId = null;
      this.userId = null;
      this.files = null;
      this.fileExtensions = null;
      this.fileListQuery = null;
      this.fileQueries = null;
      this.itemPath = null;
      this.insightType = null;
      // set the item id to load the folder
      this.itemId = item.id;
      this.fireCustomEvent('itemClick', item);
    }
  }

  private handleBreadcrumbClick(e: CustomEvent<FileListBreadCrumb>): void {
    const b = e.detail;
    this.breadcrumb = this.breadcrumb.slice(0, this.breadcrumb.indexOf(b) + 1);
    this.siteId = b.siteId;
    this.groupId = b.groupId;
    this.driveId = b.driveId;
    this.userId = b.userId;
    this.files = b.files;
    this.fileExtensions = b.fileExtensions;
    this.fileListQuery = b.fileListQuery;
    this.fileQueries = b.fileQueries;
    this.itemPath = b.itemPath;
    this.insightType = b.insightType;
    this.itemId = b.itemId;
    this.fireCustomEvent('breadcrumbclick', b);
  }
}
