import { customElement, mgtHtml, MgtTemplatedComponent } from '@microsoft/mgt-element';
import { DriveItem } from '@microsoft/microsoft-graph-types';
import { html, nothing } from 'lit';
import { property, state } from 'lit/decorators.js';
import { ifDefined } from 'lit/directives/if-defined.js';
import { OfficeGraphInsightString, ViewType } from '../../graph/types';
import { BreadcrumbInfo } from '../mgt-breadcrumb/mgt-breadcrumb';
import { strings } from './strings';

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
@customElement('file-composite')
class MgtFileComposite extends MgtTemplatedComponent {
  constructor() {
    super();
  }

  private _breadcrumbRootName: string = strings.rootNode;
  /**
   * Name to be used for the root node of the breadcrumb
   *
   * @type {string}
   * @memberof MgtFileComposite
   */
  @property({
    attribute: 'breadcrumb-root-name'
  })
  public get breadcrumbRootName(): string {
    return this._breadcrumbRootName;
  }
  public set breadcrumbRootName(value: string) {
    this._breadcrumbRootName = value;
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
  public fileListQuery: string;

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
  public fileQueries: string[];

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
  public fileExtensions: string[];

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

  /**
   * A number value indication for file size upload (KB)
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
   * @type {number}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'max-upload-file',
    type: Number
  })
  public maxUploadFile: number;

  /**
   * A Array of file extensions to be excluded from file upload.
   *
   * @type {string[]}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'excluded-file-extensions',
    converter: value => {
      return value.split(',').map(v => v.trim());
    }
  })
  public excludedFileExtensions: string[];

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
    return html`
      ${this.renderBreadcrumb()}
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
