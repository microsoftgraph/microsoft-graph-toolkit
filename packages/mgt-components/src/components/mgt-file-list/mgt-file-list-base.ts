import { arraysAreEqual } from '@microsoft/mgt-element';
import { DriveItem } from '@microsoft/microsoft-graph-types';
import { property } from 'lit/decorators.js';
import { ViewType } from '../../graph/types';
import { MgtFileBase } from '../mgt-file/mgt-file-base';

/**
 * Provides the common properties for file list components
 *
 * @abstract
 * @class MgtFileListBase
 * @extends {MgtFileBase}
 */
export abstract class MgtFileListBase extends MgtFileBase {
  private _fileListQuery: string;
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

  private _fileQueries: string[];
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

  private _fileExtensions: string[];
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
    void this.requestStateUpdate(true);
  }

  private _pageSize: number;
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

  private _maxFileSize: number;
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

  private _maxUploadFile: number;
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

  private _excludedFileExtensions: string[];
  /**
   * A Array of file extensions to be excluded from file upload.
   *
   * @type {string[]}
   * @memberof MgtFileList
   */
  @property({
    attribute: 'excluded-file-extensions',
    converter: (value, type) => {
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
}
