/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { fluentButton, fluentCheckbox, fluentDialog } from '@fluentui/web-components';
import { arraysAreEqual, customElement, IGraph, MgtBaseComponent, mgtHtml } from '@microsoft/mgt-element';
import { html, nothing, TemplateResult } from 'lit';
import { property } from 'lit/decorators.js';
import { DriveItem } from '@microsoft/microsoft-graph-types';
import {
  getGraphfile,
  getUploadSession,
  sendFileContent,
  sendFileChunk,
  deleteSessionFile,
  isUploadSession
} from '../../../graph/graph.files';
import { ViewType } from '../../../graph/types';
import { registerFluentComponents } from '../../../utils/FluentComponents';
import { getSvg, SvgIcon } from '../../../utils/SvgHelper';
import { formatBytes } from '../../../utils/Utils';
import { styles } from './mgt-file-upload-css';
import { strings } from './strings';
import './mgt-file-upload-progress';

registerFluentComponents(fluentButton, fluentCheckbox, fluentDialog);

/**
 * Simple union type for file system entry and directory entry types
 */
type FileEntry = (FileSystemDirectoryEntry | FileSystemFileEntry | FileSystemEntry) & {
  /**
   * This is a hack to get around the fact that the FileSystemEntry filePath is not writable
   */
  fullPath: string;
};

/**
 * Type guard for FileSystemDirectoryEntry
 *
 * @param {FileEntry} entry
 * @return {*}  {entry is FileSystemDirectoryEntry}
 */
const isFileSystemDirectoryEntry = (entry: FileEntry): entry is FileSystemDirectoryEntry => {
  return entry.isDirectory;
};

const isDataTransferItem = (item: DataTransferItem | File): item is DataTransferItem => {
  return (item as DataTransferItem).kind !== undefined;
};

/**
 * Type guard for FileSystemDirectoryEntry
 *
 * @param {FileEntry} entry
 * @return {*}  {entry is FileSystemDirectoryEntry}
 */
const isFileSystemFileEntry = (entry: FileEntry): entry is FileSystemFileEntry => {
  return entry.isFile;
};

interface FutureDataTransferItem extends DataTransferItem {
  /**
   * Possible future implementation of webkitGetAsEntry
   */
  getAsEntry: typeof DataTransferItem.prototype.webkitGetAsEntry;
}

/**
 * Upload conflict behavior status
 */
export const enum MgtFileUploadConflictBehavior {
  rename,
  replace
}

/**
 * MgtFileUpload upload item lifecycle object.
 *
 * @export
 * @interface MgtFileUploadItem
 */
export interface MgtFileUploadItem {
  /**
   * Session url to keep upload progress open untill all chuncks are sent
   */
  uploadUrl?: string;

  /**
   * Upload file progress value
   */
  percent?: number;

  /**
   * Validate if File has any conflict Behavior
   */
  conflictBehavior?: MgtFileUploadConflictBehavior;

  /**
   * Output "Success" or "Fail" icon base on upload response
   */
  iconStatus?: TemplateResult;

  /**
   * File object to be upload.
   */
  file?: File;

  /**
   * Full file Path to be upload.
   */
  fullPath?: string;

  /**
   * Mgt-File View state change on upload response
   */
  view?: ViewType;

  /**
   * Manipulate fileDetails on upload lifecycle
   */
  driveItem?: DriveItem;

  /**
   * Mgt-File line2Property output field message
   */
  fieldUploadResponse?: string;

  /**
   * Validates state of upload progress
   */
  completed?: boolean;

  /**
   * Load large Files into ArrayBuffer to send by chuncks
   */
  mimeStreamString?: ArrayBuffer;

  /**
   * Max chunck size to upload file by slice
   */
  maxSize?: number;

  /**
   * Minimal chunck size to upload file by slice
   */
  minSize?: number;
}

/**
 * MgtFileUpload configuration object with MgtFileList Properties.
 *
 * @export
 * @interface MgtFileUploadConfig
 */
export interface MgtFileUploadConfig {
  /**
   * MS Graph APIs connector
   *
   * @type {IGraph}
   */
  graph: IGraph;

  /**
   * allows developer to provide site id for a file
   *
   * @type {string}
   */
  siteId?: string;

  /**
   * DriveId to upload Files
   *
   * @type {string}
   */
  driveId?: string;

  /**
   * GroupId to upload Files
   *
   * @type {string}
   */
  groupId?: string;

  /**
   * allows developer to provide item id for a file
   *
   * @type {string}
   */
  itemId?: string;

  /**
   * allows developer to provide item path for a file
   *
   * @type {string}
   */
  itemPath?: string;

  /**
   * allows developer to provide user id for a file
   *
   * @type {string}
   */
  userId?: string;

  /**
   * A number value indication for file size upload (KB)
   *
   * @type {Number}
   */
  maxFileSize?: number;

  /**
   * A number value to indicate the number of files to upload.
   *
   * @type {Number}
   */
  maxUploadFile?: number;

  /**
   * A Array of file extensions to be excluded from file upload.
   *
   * @type {string[]}
   */
  excludedFileExtensions?: string[];

  /**
   * The element to use as the drop target for drag and drop.
   * Will use the parent element if not specified and available.
   *
   * @type {HTMLElement}
   * @memberof MgtFileUploadConfig
   */
  dropTarget?: () => HTMLElement;
}

interface FileWithPath extends File {
  fullPath: string;
}

/**
 * A component to upload files to OneDrive or SharePoint Sites
 *
 * @export
 * @class MgtFileUpload
 * @extends {MgtBaseComponent}
 *
 * @fires - fileUploadSuccess {undefined} - Fired when a file is successfully uploaded.
 * @fires - fileUploadChanged {MgtFileUploadItem[]} - Fired when file upload beings, changes state or is completed
 *
 * @cssprop --file-upload-background-color-drag - {Color} background color of the file list when you upload by drag and drop.
 * @cssprop --file-upload-button-background-color - {Color} background color of the file upload button.
 * @cssprop --file-upload-button-background-color-hover - {Color} background color of the file upload button on hover.
 * @cssprop --file-upload-button-text-color - {Color} text color of the file upload button.
 * @cssprop --file-upload-dialog-background-color - {Color} background color of the file upload dialog box (appears when uploaded files exist).
 * @cssprop --file-upload-dialog-text-color - {Color} text color of the file upload dialog box content.
 * @cssprop --file-upload-dialog-replace-button-background-color - {Color} background color of the replace button in the dialog box.
 * @cssprop --file-upload-dialog-replace-button-background-color-hover - {Color} background color of the replace button in the dialog box when you hover on it.
 * @cssprop --file-upload-dialog-replace-button-text-color - {Color} text color of the replace button in the dialog box.
 * @cssprop --file-upload-dialog-keep-both-button-background-color - {Color} background color of the keep-both button in the dialog box.
 * @cssprop --file-upload-dialog-keep-both-button-background-color-hover - {Color} background color of the keep-both button in the dialog box when you hover on it.
 * @cssprop --file-upload-dialog-keep-both-button-text-color - {Color} text color of the keep-both button in the dialog box.
 * @cssprop --file-upload-border-drag - {String} the border of the file list when you upload files via drag and drop. Default value is 1px dashed #0078d4.
 * @cssprop --file-upload-button-border - {String} the border of the file upload button. Default value is none.
 * @cssprop --file-upload-dialog-replace-button-border - {String} the border of the file upload replace button in the dialog box. Default value is
 * @cssprop --file-upload-dialog-keep-both-button-border - {String} the border of the file upload keep both button in the dialog box. Default value is none.
 * @cssprop --file-upload-dialog-border - {String} the border of the file upload dialog box. Default value is "1px solid var(--neutral-fill-rest)".
 * @cssprop --file-upload-dialog-width - {String} the width of the file upload dialog box. Default value is auto.
 * @cssprop --file-upload-dialog-height - {String} the height of the file upload dialog box. Default value is auto.
 * @cssprop --file-upload-dialog-padding - {String} the padding of the file upload dialog box. Default value is 24px;
 */
@customElement('file-upload')
export class MgtFileUpload extends MgtBaseComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * Strings to be used for localization
   *
   * @readonly
   * @protected
   * @memberof MgtFileUpload
   */
  protected get strings() {
    return strings;
  }

  /**
   * Disables the upload progress reporting baked into the file upload component
   *
   * @memberof MgtFileUpload
   */
  @property({
    attribute: 'hide-inline-progress',
    type: Boolean
  })
  public hideInlineProgress = false;

  private _filesToUpload: MgtFileUploadItem[] = [];
  /**
   * Allows developer to provide an array of MgtFileUploadItem to upload
   *
   * @type {MgtFileUploadItem[]}
   * @memberof MgtFileUpload
   */
  @property({ type: Object, attribute: null })
  public get filesToUpload(): MgtFileUploadItem[] {
    return this._filesToUpload;
  }
  public set filesToUpload(value: MgtFileUploadItem[]) {
    if (!arraysAreEqual(this._filesToUpload, value)) {
      this._filesToUpload = value;
      this.emitFileUploadChanged();
      void this.requestStateUpdate();
    }
  }

  private emitFileUploadChanged() {
    this.fireCustomEvent('fileUploadChanged', this._filesToUpload, true, true, true);
  }

  /**
   * List of mgt-file-list properties used to upload files.
   *
   * @type {MgtFileUploadConfig}
   * @memberof MgtFileUpload
   */
  @property({ type: Object })
  public fileUploadList: MgtFileUploadConfig;

  /**
   * Get the scopes required for file upload
   *
   * @static
   * @return {*}  {string[]}
   * @memberof MgtFileUpload
   */
  public static get requiredScopes(): string[] {
    return [...new Set(['files.readwrite', 'files.readwrite.all', 'sites.readwrite.all'])];
  }

  // variable manage drag style when mouse over
  private _dragCounter = 0;
  // variable avoids removal of files after drag and drop, https://developer.mozilla.org/en-US/docs/Web/API/DataTransfer/dropEffect
  private _dropEffect: DataTransfer['dropEffect'] = 'copy';
  // variable defined max chuck size "4MB" for large files .
  private _maxChunkSize: number = 4 * 1024 * 1024;
  private _dialogTitle = '';
  private _dialogContent = '';
  private _dialogPrimaryButton = '';
  private _dialogSecondaryButton = '';
  private _dialogCheckBox = '';
  private _applyAll = false;
  private _applyAllConflictBehavior: number = null;
  private _maximumFileSize = false;
  private _excludedFileType = false;

  private _dropTarget: HTMLElement;

  constructor() {
    super();
  }

  /**
   * Render the file upload component
   *
   * @returns
   */
  public render(): TemplateResult {
    return html`
        <div id="file-upload-dialog" class="file-upload-dialog">
          <!-- Modal content -->
          <fluent-dialog modal="true" class="file-upload-dialog-content">
            <span
              class="file-upload-dialog-close"
              id="file-upload-dialog-close">
                ${getSvg(SvgIcon.Cancel)}
            </span>
            <div class="file-upload-dialog-content-text">
              <h2 class="file-upload-dialog-title">${this._dialogTitle}</h2>
              <div>${this._dialogContent}</div>
                <fluent-checkbox
                  id="file-upload-dialog-check"
                  class="file-upload-dialog-check">
                    ${this._dialogCheckBox}
                </fluent-checkbox>
            </div>
            <div class="file-upload-dialog-editor">
              <fluent-button
                appearance="accent"
                class="file-upload-dialog-ok">
                ${this._dialogPrimaryButton}
              </fluent-button>
              <fluent-button
                appearance="outline"
                class="file-upload-dialog-cancel">
                ${this._dialogSecondaryButton}
              </fluent-button>
            </div>
          </fluent-dialog>
        </div>
        <div id="file-upload-border"></div>
        <div part="upload-button-wrapper" class="file-upload-area-button">
          <input
            id="file-upload-input"
            title="${this.strings.uploadButtonLabel}"
            tabindex="-1"
            aria-label="file upload input"
            type="file"
            multiple
            @change="${this.onFileUploadChange}"
          />
          <fluent-button
            appearance="accent"
            @click=${this.onFileUploadClick}
            label=${this.strings.uploadButtonLabel}>
              <span slot="start">${getSvg(SvgIcon.Upload)}</span>
              <span class="upload-text">${this.strings.buttonUploadFile}</span>
          </fluent-button>
        </div>
        ${
          // slice used here to create new array on each render to ensure that file-upload-progress re-renders as the data changes
          !this.hideInlineProgress
            ? mgtHtml`
              <mgt-file-upload-progress
                .progressItems=${this.filesToUpload.slice()}
                @clearnotification=${this.clearUploadNotification}
              ></mgt-file-upload-progress>`
            : nothing
        }
       `;
  }

  private clearUploadNotification = (event: CustomEvent<MgtFileUploadItem>) => {
    this.filesToUpload = this.filesToUpload.filter(item => item !== event.detail);
  };

  // TODO: remove these event listeners when component is disconnected
  // TODO: only add eventlistners we don't have them already
  public attachEventListeners() {
    const root = this.fileUploadList.dropTarget?.() || this.parentElement;
    if (root === this._dropTarget) return;
    if (root) {
      root.addEventListener('dragenter', this.handleonDragEnter);
      root.addEventListener('dragleave', this.handleonDragLeave);
      root.addEventListener('dragover', this.handleonDragOver);
      root.addEventListener('drop', this.handleOnDrop);
      this._dropTarget = root;
    }
  }

  /**
   * Handle the "Upload Files" button click event to open dialog and select files.
   *
   * @param event
   * @returns
   */
  protected onFileUploadChange = (event: UIEvent) => {
    const inputElement = event.target as HTMLInputElement;
    if (!event || inputElement.files.length < 1) {
      return;
    } else {
      void this.readUploadedFiles(inputElement.files, () => (inputElement.value = null));
    }
  };

  /**
   * Handle the click event on upload file button that open select files dialog to upload.
   *
   */
  protected onFileUploadClick = () => {
    const uploadInput: HTMLElement = this.renderRoot.querySelector('#file-upload-input');
    uploadInput.click();
  };

  /**
   * Function delete existing file upload sessions
   *
   * @param fileItem
   */
  protected async deleteFileUploadSession(fileItem: MgtFileUploadItem) {
    try {
      if (fileItem.uploadUrl !== undefined) {
        // Responses that confirm cancelation of session.
        // 404 means (The upload session was not found/The resource could not be found/)
        // 409 means The resource has changed since the caller last read it; usually an eTag mismatch
        await deleteSessionFile(this.fileUploadList.graph, fileItem.uploadUrl);
        fileItem.uploadUrl = undefined;
        fileItem.completed = true;
        this.setUploadFail(fileItem, strings.cancelUploadFile);
      } else {
        fileItem.uploadUrl = undefined;
        fileItem.completed = true;
        this.setUploadFail(fileItem, strings.cancelUploadFile);
      }
    } catch {
      fileItem.uploadUrl = undefined;
      fileItem.completed = true;
      this.setUploadFail(fileItem, strings.cancelUploadFile);
    }
  }

  /**
   * Stop listeners from onDragOver event.
   *
   * @param event
   */
  private handleonDragOver = (event: DragEvent) => {
    event.preventDefault();
    event.stopPropagation();
    if (event.dataTransfer.items && event.dataTransfer.items.length > 0) {
      event.dataTransfer.dropEffect = event.dataTransfer.dropEffect = this._dropEffect;
    }
  };

  /**
   * Stop listeners from onDragEnter event, enable drag and drop view.
   *
   * @param event
   */
  private handleonDragEnter = (event: DragEvent) => {
    event.preventDefault();
    event.stopPropagation();
    this._dragCounter++;
    if (event.dataTransfer.items && event.dataTransfer.items.length > 0) {
      event.dataTransfer.dropEffect = this._dropEffect;
      const dragFileBorder: HTMLElement = this.renderRoot.querySelector('#file-upload-border');
      dragFileBorder.classList.add('visible');
    }
  };

  /**
   * Stop listeners from ondragenter event, disable drag and drop view.
   *
   * @param event
   */
  private handleonDragLeave = (event: DragEvent) => {
    event.preventDefault();
    event.stopPropagation();

    this._dragCounter--;
    if (this._dragCounter === 0) {
      const dragFileBorder: HTMLElement = this.renderRoot.querySelector('#file-upload-border');
      dragFileBorder.classList.remove('visible');
    }
  };

  /**
   * Stop listeners from onDrop event and process files.
   *
   * @param event
   */
  private handleOnDrop = (event: DragEvent) => {
    event.preventDefault();
    event.stopPropagation();
    const done = (): void => {
      event.dataTransfer.clearData();
    };
    const dragFileBorder: HTMLElement = this.renderRoot.querySelector('#file-upload-border');
    dragFileBorder.classList.remove('visible');
    if (event.dataTransfer && event.dataTransfer.items) {
      void this.readUploadedFiles(event.dataTransfer.items, done);
    }
    this._dragCounter = 0;
  };

  private async readUploadedFiles(uploaded: DataTransferItemList | FileList, onCompleteCallback: () => void) {
    const files = await this.getFilesFromUploadArea(uploaded);
    await this.getSelectedFiles(files);
    onCompleteCallback();
  }

  /**
   * Get Files and initalize MgtFileUploadItem object life cycle to be uploaded
   *
   * @param inputFiles
   */
  private async getSelectedFiles(files: File[]) {
    let fileItems: MgtFileUploadItem[] = [];
    const fileItemsCompleted: MgtFileUploadItem[] = [];
    this._applyAll = false;
    this._applyAllConflictBehavior = null;
    this._maximumFileSize = false;
    this._excludedFileType = false;

    // Collect ongoing upload files
    this.filesToUpload.forEach(fileItem => {
      if (!fileItem.completed) {
        fileItems.push(fileItem);
      } else {
        fileItemsCompleted.push(fileItem);
      }
    });

    for (const file of files as FileWithPath[]) {
      const fullPath = file.fullPath === '' ? '/' + file.name : file.fullPath;
      if (fileItems.filter(item => item.fullPath === fullPath).length === 0) {
        // Initialize variable for File validation
        let acceptFile = true;

        // Exclude file based on max file size allowed
        if (this.fileUploadList.maxFileSize !== undefined && acceptFile) {
          if (file.size > this.fileUploadList.maxFileSize * 1024) {
            acceptFile = false;
            if (this._maximumFileSize === false) {
              const maximumFileSize: (number | true)[] = await this.getFileUploadStatus(
                file,
                fullPath,
                'MaxFileSize',
                this.fileUploadList
              );
              if (maximumFileSize !== null) {
                if (maximumFileSize[0] === 1) {
                  this._maximumFileSize = true;
                }
              }
            }
          }
        }

        // Exclude file based on File extensions
        if (this.fileUploadList.excludedFileExtensions !== undefined) {
          if (this.fileUploadList.excludedFileExtensions.length > 0 && acceptFile) {
            if (
              this.fileUploadList.excludedFileExtensions.filter(fileExtension => {
                return file.name.toLowerCase().indexOf(fileExtension.toLowerCase()) > -1;
              }).length > 0
            ) {
              acceptFile = false;
              if (this._excludedFileType === false) {
                const excludedFileType: (number | true)[] = await this.getFileUploadStatus(
                  file,
                  fullPath,
                  'ExcludedFileType',
                  this.fileUploadList
                );
                if (excludedFileType !== null) {
                  if (excludedFileType[0] === 1) {
                    this._excludedFileType = true;
                  }
                }
              }
            }
          }
        }

        // Collect accepted files
        if (acceptFile) {
          const conflictBehavior: (number | true)[] = await this.getFileUploadStatus(
            file,
            fullPath,
            'Upload',
            this.fileUploadList
          );
          let completed = false;
          if (conflictBehavior !== null) {
            if (conflictBehavior[0] === -1) {
              completed = true;
            } else {
              this._applyAll = Boolean(conflictBehavior[0]);
              this._applyAllConflictBehavior = conflictBehavior[1] ? 1 : 0;
            }
          }

          // Initialize MgtFileUploadItem Life cycle
          fileItems.push({
            file,
            driveItem: {
              name: file.name
            },
            fullPath,
            conflictBehavior: conflictBehavior !== null ? (conflictBehavior[1] ? 1 : 0) : null,
            iconStatus: null,
            percent: 1,
            view: ViewType.image,
            completed,
            maxSize: this._maxChunkSize,
            minSize: 0
          });
        }
      }
    }
    fileItems = fileItems.sort((firstFile, secondFile) => {
      return firstFile.fullPath
        .substring(0, firstFile.fullPath.lastIndexOf('/'))
        .localeCompare(secondFile.fullPath.substring(0, secondFile.fullPath.lastIndexOf('/')));
    });
    // remove completed file report image to be reuploaded.
    fileItems.forEach(fileItem => {
      if (fileItemsCompleted.filter(item => item.fullPath === fileItem.fullPath).length !== 0) {
        const index = fileItemsCompleted.findIndex(item => item.fullPath === fileItem.fullPath);
        fileItemsCompleted.splice(index, 1);
      }
    });
    fileItems.push(...fileItemsCompleted);
    this.filesToUpload = fileItems;
    // Send multiple Files to upload
    const promises = this.filesToUpload.map(fileItem => this.sendFileItemGraph(fileItem));
    await Promise.all(promises);
  }

  /**
   * Call modal dialog to replace or keep file.
   *
   * @param file
   * @returns
   */
  private async getFileUploadStatus(
    file: File,
    fullPath: string,
    DialogStatus: string,
    fileUploadList: MgtFileUploadConfig
  ) {
    const fileUploadDialog: HTMLElement = this.renderRoot.querySelector('#file-upload-dialog');

    switch (DialogStatus) {
      case 'Upload':
        const driveItem = await getGraphfile(this.fileUploadList.graph, `${this.getGrapQuery(fullPath)}?$select=id`);
        if (driveItem !== null) {
          if (this._applyAll === true) {
            return [this._applyAll, this._applyAllConflictBehavior];
          }
          fileUploadDialog.classList.add('visible');
          this._dialogTitle = strings.fileReplaceTitle;
          this._dialogContent = strings.fileReplace.replace('{FileName}', file.name);
          this._dialogCheckBox = strings.checkApplyAll;
          this._dialogPrimaryButton = strings.buttonReplace;
          this._dialogSecondaryButton = strings.buttonKeep;
          await super.requestStateUpdate(true);

          return new Promise<number[]>(resolve => {
            const fileUploadDialogClose: HTMLElement = this.renderRoot.querySelector('.file-upload-dialog-close');
            const fileUploadDialogOk: HTMLElement = this.renderRoot.querySelector('.file-upload-dialog-ok');
            const fileUploadDialogCancel: HTMLElement = this.renderRoot.querySelector('.file-upload-dialog-cancel');
            const fileUploadDialogCheck: HTMLInputElement = this.renderRoot.querySelector('#file-upload-dialog-check');
            fileUploadDialogCheck.checked = false;
            fileUploadDialogCheck.classList.remove('hide');

            // Replace File
            const onOkDialogClick = () => {
              fileUploadDialog.classList.remove('visible');
              resolve([fileUploadDialogCheck.checked ? 1 : 0, MgtFileUploadConflictBehavior.replace]);
            };

            // Rename File
            const onCancelDialogClick = () => {
              fileUploadDialog.classList.remove('visible');
              resolve([fileUploadDialogCheck.checked ? 1 : 0, MgtFileUploadConflictBehavior.rename]);
            };

            // Cancel File
            const onCloseDialogClick = () => {
              fileUploadDialog.classList.remove('visible');
              resolve([-1]);
            };

            // Remove and include event listener to validate options.
            fileUploadDialogOk.removeEventListener('click', onOkDialogClick);
            fileUploadDialogCancel.removeEventListener('click', onCancelDialogClick);
            fileUploadDialogClose.removeEventListener('click', onCloseDialogClick);
            fileUploadDialogOk.addEventListener('click', onOkDialogClick);
            fileUploadDialogCancel.addEventListener('click', onCancelDialogClick);
            fileUploadDialogClose.addEventListener('click', onCloseDialogClick);
          });
        } else {
          return null;
        }
        break;
      case 'ExcludedFileType':
        fileUploadDialog.classList.add('visible');
        this._dialogTitle = strings.fileTypeTitle;
        this._dialogContent =
          strings.fileType.replace('{FileName}', file.name) +
          ' (' +
          fileUploadList.excludedFileExtensions.join(',') +
          ')';
        this._dialogCheckBox = strings.checkAgain;
        this._dialogPrimaryButton = strings.buttonOk;
        this._dialogSecondaryButton = strings.buttonCancel;
        await super.requestStateUpdate(true);

        return new Promise<number[]>(resolve => {
          const fileUploadDialogOk: HTMLElement = this.renderRoot.querySelector('.file-upload-dialog-ok');
          const fileUploadDialogCancel: HTMLElement = this.renderRoot.querySelector('.file-upload-dialog-cancel');
          const fileUploadDialogClose: HTMLElement = this.renderRoot.querySelector('.file-upload-dialog-close');
          const fileUploadDialogCheck: HTMLInputElement = this.renderRoot.querySelector('#file-upload-dialog-check');
          fileUploadDialogCheck.checked = false;
          fileUploadDialogCheck.classList.remove('hide');

          const onOkDialogClick = () => {
            fileUploadDialog.classList.remove('visible');
            // Confirm info
            resolve([fileUploadDialogCheck.checked ? 1 : 0]);
          };

          const onCancelDialogClick = () => {
            fileUploadDialog.classList.remove('visible');
            // Cancel all
            resolve([0]);
          };

          // Remove and include event listener to validate options.
          fileUploadDialogOk.removeEventListener('click', onOkDialogClick);
          fileUploadDialogCancel.removeEventListener('click', onCancelDialogClick);
          fileUploadDialogClose.removeEventListener('click', onCancelDialogClick);
          fileUploadDialogOk.addEventListener('click', onOkDialogClick);
          fileUploadDialogCancel.addEventListener('click', onCancelDialogClick);
          fileUploadDialogClose.addEventListener('click', onCancelDialogClick);
        });
      case 'MaxFileSize':
        fileUploadDialog.classList.add('visible');
        this._dialogTitle = strings.maximumFileSizeTitle;
        this._dialogContent =
          strings.maximumFileSize
            .replace('{FileSize}', formatBytes(fileUploadList.maxFileSize * 1024))
            .replace('{FileName}', file.name) +
          formatBytes(file.size) +
          '.';
        this._dialogCheckBox = strings.checkAgain;
        this._dialogPrimaryButton = strings.buttonOk;
        this._dialogSecondaryButton = strings.buttonCancel;
        await super.requestStateUpdate(true);

        return new Promise<number[]>(resolve => {
          const fileUploadDialogOk: HTMLElement = this.renderRoot.querySelector('.file-upload-dialog-ok');
          const fileUploadDialogCancel: HTMLElement = this.renderRoot.querySelector('.file-upload-dialog-cancel');
          const fileUploadDialogClose: HTMLElement = this.renderRoot.querySelector('.file-upload-dialog-close');
          const fileUploadDialogCheck: HTMLInputElement = this.renderRoot.querySelector('#file-upload-dialog-check');
          fileUploadDialogCheck.checked = false;
          fileUploadDialogCheck.classList.remove('hide');

          const onOkDialogClick = () => {
            fileUploadDialog.classList.remove('visible');
            // Confirm info
            resolve([fileUploadDialogCheck.checked ? 1 : 0]);
          };

          const onCancelDialogClick = () => {
            fileUploadDialog.classList.remove('visible');
            // Cancel all
            resolve([0]);
          };
          // Remove and include event listener to validate options.
          fileUploadDialogOk.removeEventListener('click', onOkDialogClick);
          fileUploadDialogCancel.removeEventListener('click', onCancelDialogClick);
          fileUploadDialogClose.removeEventListener('click', onCancelDialogClick);
          fileUploadDialogOk.addEventListener('click', onOkDialogClick);
          fileUploadDialogCancel.addEventListener('click', onCancelDialogClick);
          fileUploadDialogClose.addEventListener('click', onCancelDialogClick);
        });

      default:
        break;
    }
  }

  /**
   * Get GraphQuery based on pre defined parameters.
   *
   * @param fileItem
   * @returns
   */
  private getGrapQuery(fullPath: string) {
    let itemPath = '';
    if (this.fileUploadList.itemPath) {
      if (this.fileUploadList.itemPath.length > 0) {
        itemPath =
          this.fileUploadList.itemPath.substring(0, 1) === '/'
            ? this.fileUploadList.itemPath
            : '/' + this.fileUploadList.itemPath;
      }
    }

    // {userId} {itemId}
    if (this.fileUploadList.userId && this.fileUploadList.itemId) {
      return `/users/${this.fileUploadList.userId}/drive/items/${this.fileUploadList.itemId}:${fullPath}`;
    }
    // {userId} {itemPath}
    if (this.fileUploadList.userId && this.fileUploadList.itemPath) {
      return `/users/${this.fileUploadList.userId}/drive/root:${itemPath}${fullPath}`;
    }
    // {groupId} {itemId}
    if (this.fileUploadList.groupId && this.fileUploadList.itemId) {
      return `/groups/${this.fileUploadList.groupId}/drive/items/${this.fileUploadList.itemId}:${fullPath}`;
    }
    // {groupId} {itemPath}
    if (this.fileUploadList.groupId && this.fileUploadList.itemPath) {
      return `/groups/${this.fileUploadList.groupId}/drive/root:${itemPath}${fullPath}`;
    }
    // {driveId} {itemId}
    if (this.fileUploadList.driveId && this.fileUploadList.itemId) {
      return `/drives/${this.fileUploadList.driveId}/items/${this.fileUploadList.itemId}:${fullPath}`;
    }
    // {driveId} {itemPath}
    if (this.fileUploadList.driveId && this.fileUploadList.itemPath) {
      return `/drives/${this.fileUploadList.driveId}/root:${itemPath}${fullPath}`;
    }
    // {siteId} {itemId}
    if (this.fileUploadList.siteId && this.fileUploadList.itemId) {
      return `/sites/${this.fileUploadList.siteId}/drive/items/${this.fileUploadList.itemId}:${fullPath}`;
    }
    // {siteId} {itemPath}
    if (this.fileUploadList.siteId && this.fileUploadList.itemPath) {
      return `/sites/${this.fileUploadList.siteId}/drive/root:${itemPath}${fullPath}`;
    }
    // default OneDrive {itemId}
    if (this.fileUploadList.itemId) {
      return `/me/drive/items/${this.fileUploadList.itemId}:${fullPath}`;
    }
    // default OneDrive {itemPath}
    if (this.fileUploadList.itemPath) {
      return `/me/drive/root:${itemPath}${fullPath}`;
    }
    // default OneDrive root
    return `/me/drive/root:${fullPath}`;
  }

  /**
   * Send file using Upload using Graph based on length
   *
   * @param fileUpload
   * @returns
   */
  private async sendFileItemGraph(fileItem: MgtFileUploadItem) {
    const graph: IGraph = this.fileUploadList.graph;
    let graphQuery = '';
    if (fileItem.file.size < this._maxChunkSize) {
      try {
        if (!fileItem.completed) {
          if (
            fileItem.conflictBehavior === null ||
            fileItem.conflictBehavior === MgtFileUploadConflictBehavior.replace
          ) {
            graphQuery = `${this.getGrapQuery(fileItem.fullPath)}:/content`;
          }
          if (fileItem.conflictBehavior === MgtFileUploadConflictBehavior.rename) {
            graphQuery = `${this.getGrapQuery(fileItem.fullPath)}:/content?@microsoft.graph.conflictBehavior=rename`;
          }
          fileItem.driveItem = await sendFileContent(graph, graphQuery, fileItem.file);
          if (fileItem.driveItem !== null) {
            this.setUploadSuccess(fileItem);
          } else {
            fileItem.driveItem = {
              name: fileItem.file.name
            };
            this.setUploadFail(fileItem, strings.failUploadFile);
          }
        }
      } catch (error) {
        this.setUploadFail(fileItem, strings.failUploadFile);
      }
    } else {
      if (!fileItem.completed) {
        if (fileItem.uploadUrl === undefined) {
          const response = await getUploadSession(
            graph,
            `${this.getGrapQuery(fileItem.fullPath)}:/createUploadSession`,
            fileItem.conflictBehavior
          );
          try {
            if (response !== null) {
              // uploadSession url used to send chunks of file
              fileItem.uploadUrl = response.uploadUrl;
              const driveItem = await this.sendSessionUrlGraph(graph, fileItem);
              if (driveItem !== null) {
                fileItem.driveItem = driveItem;
                this.setUploadSuccess(fileItem);
              } else {
                this.setUploadFail(fileItem, strings.failUploadFile);
              }
            } else {
              this.setUploadFail(fileItem, strings.failUploadFile);
            }
            // eslint-disable-next-line no-empty
          } catch {}
        }
      }
    }
  }

  /**
   * Manage slices of File to upload file by chunks using Graph and Session Url
   *
   * @param Graph
   * @param fileItem
   * @returns
   */
  private async sendSessionUrlGraph(graph: IGraph, fileItem: MgtFileUploadItem) {
    while (fileItem.file.size > fileItem.minSize) {
      if (fileItem.mimeStreamString === undefined) {
        fileItem.mimeStreamString = (await this.readFileContent(fileItem.file)) as ArrayBuffer;
      }
      // Graph client API uses Buffer package to manage ArrayBuffer, change to Blob avoids external package dependency
      const fileSlice: Blob = new Blob([fileItem.mimeStreamString.slice(fileItem.minSize, fileItem.maxSize)]);
      fileItem.percent = Math.round((fileItem.maxSize / fileItem.file.size) * 100);
      await super.requestStateUpdate(true);
      // emit update here as the percent of the upload for the file is changed.
      this.emitFileUploadChanged();

      if (fileItem.uploadUrl !== undefined) {
        const response = await sendFileChunk(
          graph,
          fileItem.uploadUrl,
          `${fileItem.maxSize - fileItem.minSize}`,
          `bytes ${fileItem.minSize}-${fileItem.maxSize - 1}/${fileItem.file.size}`,
          fileSlice
        );
        if (response === null) {
          return null;
        } else if (isUploadSession(response)) {
          // Define next Chunk
          fileItem.minSize = parseInt(response.nextExpectedRanges[0].split('-')[0], 10);
          fileItem.maxSize = fileItem.minSize + this._maxChunkSize;
          if (fileItem.maxSize > fileItem.file.size) {
            fileItem.maxSize = fileItem.file.size;
          }
        } else if (response.id !== undefined) {
          return response;
        }
      } else {
        return null;
      }
    }
  }

  /**
   * Change the state of Mgt-File icon upload to Success
   *
   * @param fileUpload
   */
  private setUploadSuccess(fileUpload: MgtFileUploadItem) {
    fileUpload.percent = 100;
    fileUpload.iconStatus = getSvg(SvgIcon.Success);
    fileUpload.view = ViewType.twolines;
    fileUpload.fieldUploadResponse = 'lastModifiedDateTime';
    fileUpload.completed = true;
    this.requestUpdate();
    this.emitFileUploadChanged();
    this.fireCustomEvent('fileUploadSuccess', undefined, true, true, true);
  }

  /**
   * Change the state of Mgt-File icon upload to Fail
   *
   * @param fileUpload
   */
  private setUploadFail(fileUpload: MgtFileUploadItem, errorMessage: string) {
    fileUpload.iconStatus = getSvg(SvgIcon.Fail);
    fileUpload.view = ViewType.twolines;
    fileUpload.driveItem.description = errorMessage;
    fileUpload.fieldUploadResponse = 'description';
    fileUpload.completed = true;
    this.emitFileUploadChanged();
    super.requestUpdate();
  }

  /**
   * Retrieve File content as ArrayBuffer
   *
   * @param file
   * @returns
   */
  private readFileContent(file: File): Promise<string | ArrayBuffer> {
    return new Promise<string | ArrayBuffer>((resolve, reject) => {
      const myReader: FileReader = new FileReader();

      myReader.onloadend = () => {
        resolve(myReader.result);
      };

      myReader.onerror = e => {
        reject(e);
      };

      myReader.readAsArrayBuffer(file);
    });
  }

  /**
   * Collect Files from Upload Area based on maxUploadFile
   *
   * @param uploadFilesItems
   * @returns
   */
  protected async getFilesFromUploadArea(filesItems: DataTransferItemList | FileList): Promise<File[]> {
    const folders: FileSystemDirectoryEntry[] = [];
    let entry: FileSystemEntry;
    const collectFilesItems: File[] = [];

    for (const uploadFileItem of filesItems) {
      if (isDataTransferItem(uploadFileItem) && uploadFileItem.kind === 'file') {
        // Defensive code to validate if function exists in Browser
        // Collect all Folders into Array
        const futureUpload = uploadFileItem as FutureDataTransferItem;
        if (futureUpload.getAsEntry) {
          entry = futureUpload.getAsEntry();
          if (isFileSystemDirectoryEntry(entry)) {
            folders.push(entry);
          } else {
            const file = uploadFileItem.getAsFile();
            if (file) {
              this.writeFilePath(file, '');
              collectFilesItems.push(file);
            }
          }
        } else if (uploadFileItem.webkitGetAsEntry) {
          entry = uploadFileItem.webkitGetAsEntry();
          if (isFileSystemDirectoryEntry(entry)) {
            folders.push(entry);
          } else {
            const file = uploadFileItem.getAsFile();
            if (file) {
              this.writeFilePath(file, '');
              collectFilesItems.push(file);
            }
          }
        } else if ('function' === typeof uploadFileItem.getAsFile) {
          const file = uploadFileItem.getAsFile();
          if (file) {
            this.writeFilePath(file, '');
            collectFilesItems.push(file);
          }
        }
        continue;
      } else {
        const fileItem = isDataTransferItem(uploadFileItem) ? uploadFileItem.getAsFile() : uploadFileItem;
        if (fileItem) {
          this.writeFilePath(fileItem, '');
          collectFilesItems.push(fileItem);
        }
      }
    }
    // Collect Files from folder
    if (folders.length > 0) {
      const folderFiles = await this.getFolderFiles(folders);
      collectFilesItems.push(...folderFiles);
    }
    return collectFilesItems;
  }

  /**
   * Retrieve files from folder and subfolders to array.
   *
   * @param folders
   * @returns
   */
  private getFolderFiles(folders: FileSystemDirectoryEntry[]) {
    return new Promise<File[]>(resolve => {
      let reading = 0;
      const contents: File[] = [];

      const readEntry = (entry: FileEntry, path: string) => {
        if (isFileSystemDirectoryEntry(entry)) {
          readReaderContent(entry.createReader());
        } else if (isFileSystemFileEntry(entry)) {
          reading++;
          entry.file(file => {
            reading--;
            // Include Folder path where File is located
            this.writeFilePath(file, path);
            contents.push(file);

            if (reading === 0) {
              resolve(contents);
            }
          });
        }
      };

      const readReaderContent = (reader: FileSystemDirectoryReader) => {
        reading++;

        reader.readEntries(entries => {
          reading--;
          for (const entry of entries) {
            readEntry(entry, entry.fullPath);
          }

          if (reading === 0) {
            resolve(contents);
          }
        });
      };

      folders.forEach(entry => {
        readEntry(entry, '');
      });
    });
  }
  private writeFilePath(file: File | FileSystemEntry, path: string) {
    ((file as unknown) as FileEntry).fullPath = path;
  }
}
