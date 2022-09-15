/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, property, TemplateResult } from 'lit-element';
import { styles } from './mgt-file-upload-css';
import { strings } from './strings';
import { getSvg, SvgIcon } from '../../../utils/SvgHelper';
import { formatBytes } from '../../../utils/Utils';
import { IGraph, MgtBaseComponent } from '@microsoft/mgt-element';
import { ViewType } from '../../../graph/types';
import { DriveItem } from '@microsoft/microsoft-graph-types';
import {
  clearFilesCache,
  getGraphfile,
  getUploadSession,
  sendFileContent,
  sendFileChunck,
  deleteSessionFile
} from '../../../graph/graph.files';

export { FluentProgress, FluentButton, FluentCheckbox, FluentCard } from '@fluentui/web-components';

// import { registerFluentComponents } from '../../../utils/FluentComponents';
// import { fluentButton, fluentCheckbox, fluentProgress } from '@fluentui/web-components';

// registerFluentComponents(fluentProgress, fluentButton, fluentCheckbox);

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
   *  Upload file progress value
   */
  percent?: number;

  /**
   *  Validate if File has any conflict Behavior
   */
  conflictBehavior?: MgtFileUploadConflictBehavior;

  /**
   *  Output "Success" or "Fail" icon base on upload response
   */
  iconStatus?: TemplateResult;

  /**
   * File object to be upload.
   */
  file?: File;

  /**
   *  Full file Path to be upload.
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
   *  Max chunck size to upload file by slice
   */
  maxSize?: number;

  /**
   *  Minimal chunck size to upload file by slice
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
   * @type {IGraph}
   */
  graph: IGraph;

  /**
   *  allows developer to provide site id for a file
   *  @type {string}
   */
  siteId?: string;

  /**
   * DriveId to upload Files
   *  @type {string}
   */
  driveId?: string;

  /**
   * GroupId to upload Files
   *  @type {string}
   */
  groupId?: string;

  /**
   * allows developer to provide item id for a file
   *  @type {string}
   */
  itemId?: string;

  /**
   *  allows developer to provide item path for a file
   *  @type {string}
   */
  itemPath?: string;

  /**
   * allows developer to provide user id for a file
   *  @type {string}
   */
  userId?: string;

  /**
   * A number value indication for file size upload (KB)
   *  @type {Number}
   */
  maxFileSize?: number;

  /**
   *  A number value to indicate the number of files to upload.
   *  @type {Number}
   */
  maxUploadFile?: Number;

  /**
   * A Array of file extensions to be excluded from file upload.
   *
   *  @type {string[]}
   */
  excludedFileExtensions?: string[];
}

/**
 * A component to upload files to OneDrive or SharePoint Sites
 *
 * @export
 * @class MgtFileUpload
 * @extends {MgtBaseComponent}
 *
 * @cssprop --file-upload-border- {String} File upload border top style
 * @cssprop --file-upload-background-color - {Color} File upload background color with opacity style
 * @cssprop --file-upload-button-float - {string} Upload button float position
 * @cssprop --file-upload-button-background-color - {Color} Background color of upload button
 * @cssprop --file-upload-dialog-background-color - {Color} Background color of dialog
 * @cssprop --file-upload-dialog-content-background-color - {Color} Background color of dialog content
 * @cssprop --file-upload-dialog-content-color - {Color} Color of dialog content
 * @cssprop --file-upload-button-color - {Color} Text color of upload button
 * @cssprop --file-upload-dialog-primarybutton-background-color - {Color} Background color of primary button
 * @cssprop --file-upload-dialog-primarybutton-color - {Color} Color text of primary button
 * @cssprop --file-item-margin - {String} File item margin
 * @cssprop --file-item-background-color--hover - {Color} File item background hover color
 * @cssprop --file-item-border-top - {String} File item border top style
 * @cssprop --file-item-border-left - {String} File item border left style
 * @cssprop --file-item-border-right - {String} File item border right style
 * @cssprop --file-item-border-bottom - {String} File item border bottom style
 * @cssprop --file-item-background-color--active - {Color} File item background active color
 */
@customElement('mgt-file-upload')
export class MgtFileUpload extends MgtBaseComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  protected get strings() {
    return strings;
  }

  /**
   * Allows developer to provide an array of MgtFileUploadItem to upload
   *
   * @type {MgtFileUploadItem[]}
   * @memberof MgtFileUpload
   */
  @property({ type: Object })
  public filesToUpload: MgtFileUploadItem[];

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
  private _dragCounter: number = 0;
  // variable avoids removal of files after drag and drop, https://developer.mozilla.org/en-US/docs/Web/API/DataTransfer/dropEffect
  private _dropEffect: string = 'copy';
  // variable defined max chuck size "4MB" for large files .
  private _maxChunckSize: number = 4 * 1024 * 1024;
  private _dialogTitle: string = '';
  private _dialogContent: string = '';
  private _dialogPrimaryButton: string = '';
  private _dialogSecondaryButton: string = '';
  private _dialogCheckBox: string = '';
  private _applyAll: boolean = false;
  private _applyAllConflitBehavior: number = null;
  private _maximumFiles: boolean = false;
  private _maximumFileSize: boolean = false;
  private _excludedFileType: boolean = false;

  constructor() {
    super();
    this.filesToUpload = [];
  }

  /**
   *
   * @returns
   */
  public render(): TemplateResult {
    if (this.parentElement !== null) {
      const root = this.parentElement;
      root.addEventListener('dragenter', this.handleonDragEnter);
      root.addEventListener('dragleave', this.handleonDragLeave);
      root.addEventListener('dragover', this.handleonDragOver);
      root.addEventListener('drop', this.handleonDrop);
    }

    return html`
        <div id="file-upload-dialog" class="file-upload-dialog">
          <!-- Modal content -->
          <fluent-card class="file-upload-dialog-content">
            <span class="file-upload-dialog-close" id="file-upload-dialog-close" >${getSvg(SvgIcon.Cancel)}</span>
            <div class="file-upload-dialog-content-text">
              <h2 class="file-upload-dialog-title">${this._dialogTitle}</h2>
              <div>${this._dialogContent}</div>
              <div class="file-upload-dialog-check-wrapper">
                <fluent-checkbox id="file-upload-dialog-check" class="file-upload-dialog-check" >
                  <span>${this._dialogCheckBox}</span>
                </fluent-checkbox>
              </div>
            </div>
            <div class="file-upload-dialog-editor">
              <fluent-button class="file-upload-dialog-ok">
              ${this._dialogPrimaryButton}
              </fluent-button>
              <fluent-button class="file-upload-dialog-cancel">
              ${this._dialogSecondaryButton}
              </fluent-button>
            </div>
          </fluent-card>
        </div>
        <div id="file-upload-border" >
        </div>
        <div class="file-upload-area-button">
        <div>
          <input
            id="file-upload-input"
            aria-label="file upload input"
            type="file"
            multiple="true"
            @change="${this.onFileUploadChange}"
          />
          <fluent-button class="file-upload-button" @click=${this.onFileUploadClick}>
            ${getSvg(SvgIcon.Upload)}${strings.buttonUploadFile}
          </fluent-button>
        </div>
        </div>
        <div class="file-upload-Template">
        ${this.renderFolderTemplate(this.filesToUpload)}
        </div>
       `;
  }

  /**
   * Render Folder structure of files to upload
   * @param fileItems
   * @returns
   */
  protected renderFolderTemplate(fileItems: MgtFileUploadItem[]) {
    let folderStructure: string[] = [];
    if (fileItems.length > 0) {
      const TemplateFileItems = fileItems.map(fileItem => {
        if (folderStructure.indexOf(fileItem.fullPath.substring(0, fileItem.fullPath.lastIndexOf('/'))) === -1) {
          if (fileItem.fullPath.substring(0, fileItem.fullPath.lastIndexOf('/')) !== '') {
            folderStructure.push(fileItem.fullPath.substring(0, fileItem.fullPath.lastIndexOf('/')));
            return html`
            <div class='file-upload-table'>
              <div class='file-upload-cell'>
                <mgt-file
                  .fileDetails=${{
                    name: fileItem.fullPath.substring(1, fileItem.fullPath.lastIndexOf('/')),
                    folder: 'Folder'
                  }}
                  .view=${ViewType.oneline}
                  class="mgt-file-item"
                >
                </mgt-file>
              </div>
            </div>
            ${this.renderFileTemplate(fileItem, 'file-upload-folder-tab')}`;
          } else {
            return html`${this.renderFileTemplate(fileItem, '')}`;
          }
        } else {
          return html`${this.renderFileTemplate(fileItem, 'file-upload-folder-tab')}`;
        }
      });
      return html`${TemplateFileItems}`;
    } else {
      return null;
    }
  }

  /**
   * Render file upload area
   *
   * @param fileItem
   * @returns
   */
  protected renderFileTemplate(fileItem: MgtFileUploadItem, folderTabStyle: string) {
    return html`
        <div class="${fileItem.completed ? 'file-upload-table' : 'file-upload-table upload'}">
          <div class="${
            folderTabStyle +
            (fileItem.fieldUploadResponse === 'lastModifiedDateTime' ? ' file-upload-dialog-success' : '')
          }">
            <div class='file-upload-cell'>
              <div style=${fileItem.fieldUploadResponse === 'description' ? 'opacity: 0.5;' : ''}>
                <div class="file-upload-status">
                  ${fileItem.iconStatus}
                </div>
                <mgt-file
                  .fileDetails=${fileItem.driveItem}
                  .view=${fileItem.view}
                  .line2Property=${fileItem.fieldUploadResponse}
                  class="mgt-file-item"
                  >
                </mgt-file>
              </div>
            </div>
              ${fileItem.completed === false ? this.renderFileUploadTemplate(fileItem) : null}
            </div>
          </div>
        </div>
        `;
  }

  /**
   * Render file upload progress
   *
   * @param fileItem
   * @returns
   */
  protected renderFileUploadTemplate(fileItem: MgtFileUploadItem) {
    return html`
    <div class='file-upload-cell'>
      <div class='file-upload-table file-upload-name' >
        <div class='file-upload-cell'>
          <div title="${fileItem.file.name}" class='file-upload-filename'>${fileItem.file.name}</div>
        </div>
      </div>
      <div class='file-upload-table'>
        <div class='file-upload-cell'>
          <div class="${fileItem.completed ? 'file-upload-table' : 'file-upload-table upload'}">
            <fluent-progress class="file-upload-bar" value="${fileItem.percent}" ></fluent-progress>
            <div class='file-upload-cell' style="padding-left:5px">
              <span>${fileItem.percent}%</span>
              <span
                class="file-upload-cancel"
                @click=${e => this.deleteFileUploadSession(fileItem)}>
                ${getSvg(SvgIcon.Cancel)}
              </span>
            </div>
          <div>
        </div>
      </div>
    </div>
    `;
  }

  /**
   * Handle the "Upload Files" button click event to open dialog and select files.
   *
   * @param event
   * @returns
   */
  protected async onFileUploadChange(event) {
    if (!event || event.target.files.length < 1) {
      return;
    } else {
      this.getSelectedFiles(await this.getFilesFromUploadArea(event.target.files));
      event.target.value = null;
    }
  }

  /**
   * Handle the click event on upload file button that open select files dialog to upload.
   *
   */
  protected onFileUploadClick() {
    const uploadInput: HTMLElement = this.renderRoot.querySelector('#file-upload-input');
    uploadInput.click();
  }

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
        const response = await deleteSessionFile(this.fileUploadList.graph, fileItem.uploadUrl);
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
  protected handleonDragOver = async event => {
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
  protected handleonDragEnter = async event => {
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
  protected handleonDragLeave = event => {
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
  protected handleonDrop = async event => {
    event.preventDefault();
    event.stopPropagation();

    const dragFileBorder: HTMLElement = this.renderRoot.querySelector('#file-upload-border');
    dragFileBorder.classList.remove('visible');
    if (event.dataTransfer && event.dataTransfer.items) {
      this.getSelectedFiles(await this.getFilesFromUploadArea(event.dataTransfer.items));
    }
    event.dataTransfer.clearData();
    this._dragCounter = 0;
  };

  /**
   * Get Files and initalize MgtFileUploadItem object life cycle to be uploaded
   *
   * @param inputFiles
   */
  protected async getSelectedFiles(files: File[]) {
    let fileItems: MgtFileUploadItem[] = [];
    let fileItemsCompleted: MgtFileUploadItem[] = [];
    this._applyAll = false;
    this._applyAllConflitBehavior = null;
    this._maximumFiles = false;
    this._maximumFileSize = false;
    this._excludedFileType = false;

    //Collect ongoing upload files
    this.filesToUpload.forEach(async fileItem => {
      if (!fileItem.completed) {
        fileItems.push(fileItem);
      } else {
        fileItemsCompleted.push(fileItem);
      }
    });

    for (var i = 0; i < files.length; i++) {
      const file: any = files[i];
      const fullPath = file.fullPath === '' ? '/' + file.name : file.fullPath;
      if (fileItems.filter(item => item.fullPath === fullPath).length === 0) {
        //Initialize variable for File validation
        let acceptFile = true;

        //Exclude file based on max files upload allowed
        if (fileItems.length >= this.fileUploadList.maxUploadFile) {
          acceptFile = false;
          if (!this._maximumFiles) {
            const maximumFiles: (number | true)[] = await this.getFileUploadStatus(
              files[i],
              fullPath,
              'MaxFiles',
              this.fileUploadList
            );
            if (maximumFiles !== null) {
              if (maximumFiles[0] === 0) {
                return null;
              }
              if (maximumFiles[0] === 1) {
                this._maximumFiles = true;
              }
            }
          }
        }

        //Exclude file based on max file size allowed
        if (this.fileUploadList.maxFileSize !== undefined && acceptFile) {
          if (files[i].size > this.fileUploadList.maxFileSize * 1024) {
            acceptFile = false;
            if (this._maximumFileSize === false) {
              const maximumFileSize: (number | true)[] = await this.getFileUploadStatus(
                files[i],
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

        //Exclude file based on File extensions
        if (this.fileUploadList.excludedFileExtensions !== undefined) {
          if (this.fileUploadList.excludedFileExtensions.length > 0 && acceptFile) {
            if (
              this.fileUploadList.excludedFileExtensions.filter(fileExtension => {
                return files[i].name.toLowerCase().indexOf(fileExtension.toLowerCase()) > -1;
              }).length > 0
            ) {
              acceptFile = false;
              if (this._excludedFileType === false) {
                const excludedFileType: (number | true)[] = await this.getFileUploadStatus(
                  files[i],
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

        //Collect accepted files
        if (acceptFile) {
          const conflictBehavior: (number | true)[] = await this.getFileUploadStatus(
            files[i],
            fullPath,
            'Upload',
            this.fileUploadList
          );
          let _completed = false;
          if (conflictBehavior !== null) {
            if (conflictBehavior[0] === -1) {
              _completed = true;
            } else {
              this._applyAll = Boolean(conflictBehavior[0]);
              this._applyAllConflitBehavior = conflictBehavior[1] ? 1 : 0;
            }
          }

          //Initialize MgtFileUploadItem Life cycle
          fileItems.push({
            file: files[i],
            driveItem: {
              name: files[i].name
            },
            fullPath: fullPath,
            conflictBehavior: conflictBehavior !== null ? (conflictBehavior[1] ? 1 : 0) : null,
            iconStatus: null,
            percent: 1,
            view: ViewType.image,
            completed: _completed,
            maxSize: this._maxChunckSize,
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
    //remove completed file report image to be reuploaded.
    fileItems.forEach(fileItem => {
      if (fileItemsCompleted.filter(item => item.fullPath === fileItem.fullPath).length !== 0) {
        let index = fileItemsCompleted.findIndex(item => item.fullPath === fileItem.fullPath);
        fileItemsCompleted.splice(index, 1);
      }
    });
    fileItems.push(...fileItemsCompleted);
    this.filesToUpload = fileItems;
    // Send multiple Files to upload
    this.filesToUpload.forEach(async fileItem => {
      await this.sendFileItemGraph(fileItem);
    });
  }

  /**
   * Call modal dialog to replace or keep file.
   *
   * @param file
   * @returns
   */
  protected async getFileUploadStatus(
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
            return [this._applyAll, this._applyAllConflitBehavior];
          }
          fileUploadDialog.classList.add('visible');
          this._dialogTitle = strings.fileReplaceTitle;
          this._dialogContent = strings.fileReplace.replace('{FileName}', file.name);
          this._dialogCheckBox = strings.checkApplyAll;
          this._dialogPrimaryButton = strings.buttonReplace;
          this._dialogSecondaryButton = strings.buttonKeep;
          super.requestStateUpdate(true);

          return new Promise<number[]>(async resolve => {
            let fileUploadDialogClose: HTMLElement = this.renderRoot.querySelector('.file-upload-dialog-close');
            let fileUploadDialogOk: HTMLElement = this.renderRoot.querySelector('.file-upload-dialog-ok');
            let fileUploadDialogCancel: HTMLElement = this.renderRoot.querySelector('.file-upload-dialog-cancel');
            let fileUploadDialogCheck: HTMLInputElement = this.renderRoot.querySelector('#file-upload-dialog-check');
            fileUploadDialogCheck.checked = false;
            fileUploadDialogCheck.classList.remove('hide');

            //Remove and include event listener to validate options.
            fileUploadDialogOk.removeEventListener('click', onOkDialogClick);
            fileUploadDialogCancel.removeEventListener('click', onCancelDialogClick);
            fileUploadDialogClose.removeEventListener('click', onCloseDialogClick);
            fileUploadDialogOk.addEventListener('click', onOkDialogClick);
            fileUploadDialogCancel.addEventListener('click', onCancelDialogClick);
            fileUploadDialogClose.addEventListener('click', onCloseDialogClick);

            //Replace File
            function onOkDialogClick() {
              fileUploadDialog.classList.remove('visible');
              resolve([fileUploadDialogCheck.checked ? 1 : 0, MgtFileUploadConflictBehavior.replace]);
            }

            //Rename File
            function onCancelDialogClick() {
              fileUploadDialog.classList.remove('visible');
              resolve([fileUploadDialogCheck.checked ? 1 : 0, MgtFileUploadConflictBehavior.rename]);
            }

            //Cancel File
            function onCloseDialogClick() {
              fileUploadDialog.classList.remove('visible');
              resolve([-1]);
            }
          });
        } else {
          return null;
        }
        break;
      case 'MaxFiles':
        fileUploadDialog.classList.add('visible');
        this._dialogTitle = strings.maximumFilesTitle;
        this._dialogContent = strings.maximumFiles.split('{MaxNumber}').join(fileUploadList.maxUploadFile.toString());
        this._dialogCheckBox = strings.checkApplyAll;
        this._dialogPrimaryButton = strings.buttonUpload;
        this._dialogSecondaryButton = strings.buttonReselect;
        super.requestStateUpdate(true);

        return new Promise<number[]>(async resolve => {
          let fileUploadDialogOk: HTMLElement = this.renderRoot.querySelector('.file-upload-dialog-ok');
          let fileUploadDialogCancel: HTMLElement = this.renderRoot.querySelector('.file-upload-dialog-cancel');
          let fileUploadDialogClose: HTMLElement = this.renderRoot.querySelector('.file-upload-dialog-close');
          let fileUploadDialogCheck: HTMLInputElement = this.renderRoot.querySelector('#file-upload-dialog-check');
          fileUploadDialogCheck.checked = false;
          fileUploadDialogCheck.classList.add('hide');

          //Remove and include event listener to validate options.
          fileUploadDialogOk.removeEventListener('click', onOkDialogClick);
          fileUploadDialogCancel.removeEventListener('click', onCancelDialogClick);
          fileUploadDialogClose.removeEventListener('click', onCancelDialogClick);
          fileUploadDialogOk.addEventListener('click', onOkDialogClick);
          fileUploadDialogCancel.addEventListener('click', onCancelDialogClick);
          fileUploadDialogClose.addEventListener('click', onCancelDialogClick);

          function onOkDialogClick() {
            fileUploadDialog.classList.remove('visible');
            //Continue upload
            resolve([1]);
          }

          function onCancelDialogClick() {
            fileUploadDialog.classList.remove('visible');
            //Cancel all
            resolve([0]);
          }
        });
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
        super.requestStateUpdate(true);

        return new Promise<number[]>(async resolve => {
          let fileUploadDialogOk: HTMLElement = this.renderRoot.querySelector('.file-upload-dialog-ok');
          let fileUploadDialogCancel: HTMLElement = this.renderRoot.querySelector('.file-upload-dialog-cancel');
          let fileUploadDialogClose: HTMLElement = this.renderRoot.querySelector('.file-upload-dialog-close');
          let fileUploadDialogCheck: HTMLInputElement = this.renderRoot.querySelector('#file-upload-dialog-check');
          fileUploadDialogCheck.checked = false;
          fileUploadDialogCheck.classList.remove('hide');

          //Remove and include event listener to validate options.
          fileUploadDialogOk.removeEventListener('click', onOkDialogClick);
          fileUploadDialogCancel.removeEventListener('click', onCancelDialogClick);
          fileUploadDialogClose.removeEventListener('click', onCancelDialogClick);
          fileUploadDialogOk.addEventListener('click', onOkDialogClick);
          fileUploadDialogCancel.addEventListener('click', onCancelDialogClick);
          fileUploadDialogClose.addEventListener('click', onCancelDialogClick);

          function onOkDialogClick() {
            fileUploadDialog.classList.remove('visible');
            //Confirm info
            resolve([fileUploadDialogCheck.checked ? 1 : 0]);
          }

          function onCancelDialogClick() {
            fileUploadDialog.classList.remove('visible');
            //Cancel all
            resolve([0]);
          }
        });
      case 'MaxFileSize':
        fileUploadDialog.classList.add('visible');
        this._dialogTitle = strings.maximumFileSizeTitle;
        this._dialogContent =
          strings.maximumFileSize
            .replace('{FileSize}', formatBytes(fileUploadList.maxFileSize))
            .replace('{FileName}', file.name) +
          formatBytes(file.size) +
          '.';
        this._dialogCheckBox = strings.checkAgain;
        this._dialogPrimaryButton = strings.buttonOk;
        this._dialogSecondaryButton = strings.buttonCancel;
        super.requestStateUpdate(true);

        return new Promise<number[]>(async resolve => {
          let fileUploadDialogOk: HTMLElement = this.renderRoot.querySelector('.file-upload-dialog-ok');
          let fileUploadDialogCancel: HTMLElement = this.renderRoot.querySelector('.file-upload-dialog-cancel');
          let fileUploadDialogClose: HTMLElement = this.renderRoot.querySelector('.file-upload-dialog-close');
          let fileUploadDialogCheck: HTMLInputElement = this.renderRoot.querySelector('#file-upload-dialog-check');
          fileUploadDialogCheck.checked = false;
          fileUploadDialogCheck.classList.remove('hide');

          //Remove and include event listener to validate options.
          fileUploadDialogOk.removeEventListener('click', onOkDialogClick);
          fileUploadDialogCancel.removeEventListener('click', onCancelDialogClick);
          fileUploadDialogClose.removeEventListener('click', onCancelDialogClick);
          fileUploadDialogOk.addEventListener('click', onOkDialogClick);
          fileUploadDialogCancel.addEventListener('click', onCancelDialogClick);
          fileUploadDialogClose.addEventListener('click', onCancelDialogClick);

          function onOkDialogClick() {
            fileUploadDialog.classList.remove('visible');
            //Confirm info
            resolve([fileUploadDialogCheck.checked ? 1 : 0]);
          }

          function onCancelDialogClick() {
            fileUploadDialog.classList.remove('visible');
            //Cancel all
            resolve([0]);
          }
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
  protected getGrapQuery(fullPath: string) {
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
  protected async sendFileItemGraph(fileItem: MgtFileUploadItem) {
    const graph: IGraph = this.fileUploadList.graph;
    let graphQuery = '';
    if (fileItem.file.size < this._maxChunckSize) {
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
              // uploadSession url used to send chuncks of file
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
          } catch {}
        }
      }
    }
  }

  /**
   * Manage slices of File to upload file by chuncks using Graph and Session Url
   *
   * @param Graph
   * @param fileItem
   * @returns
   */
  protected async sendSessionUrlGraph(graph: IGraph, fileItem: MgtFileUploadItem) {
    while (fileItem.file.size > fileItem.minSize) {
      if (fileItem.mimeStreamString === undefined) {
        fileItem.mimeStreamString = (await this.readFileContent(fileItem.file)) as ArrayBuffer;
      }
      //Graph client API uses Buffer package to manage ArrayBuffer, change to Blob avoids external package dependency
      const fileSlice: Blob = new Blob([fileItem.mimeStreamString.slice(fileItem.minSize, fileItem.maxSize)]);
      fileItem.percent = Math.round((fileItem.maxSize / fileItem.file.size) * 100);
      super.requestStateUpdate(true);

      if (fileItem.uploadUrl !== undefined) {
        const response = await sendFileChunck(
          graph,
          fileItem.uploadUrl,
          `${fileItem.maxSize - fileItem.minSize}`,
          `bytes ${fileItem.minSize}-${fileItem.maxSize - 1}/${fileItem.file.size}`,
          fileSlice
        );
        if (response === null) {
          return null;
        } else if (response.id !== undefined) {
          return response as DriveItem;
        } else if (response.nextExpectedRanges !== undefined) {
          //Define next Chunck
          fileItem.minSize = parseInt(response.nextExpectedRanges[0].split('-')[0]);
          fileItem.maxSize = fileItem.minSize + this._maxChunckSize;
          if (fileItem.maxSize > fileItem.file.size) {
            fileItem.maxSize = fileItem.file.size;
          }
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
  protected setUploadSuccess(fileUpload: MgtFileUploadItem) {
    fileUpload.percent = 100;
    super.requestStateUpdate(true);
    setTimeout(() => {
      fileUpload.iconStatus = getSvg(SvgIcon.Success);
      fileUpload.view = ViewType.twolines;
      fileUpload.fieldUploadResponse = 'lastModifiedDateTime';
      fileUpload.completed = true;
      super.requestStateUpdate(true);
      clearFilesCache();
    }, 500);
  }

  /**
   * Change the state of Mgt-File icon upload to Fail
   *
   * @param fileUpload
   */
  protected setUploadFail(fileUpload: MgtFileUploadItem, errorMessage: string) {
    setTimeout(() => {
      fileUpload.iconStatus = getSvg(SvgIcon.Fail);
      fileUpload.view = ViewType.twolines;
      fileUpload.driveItem.description = errorMessage;
      fileUpload.fieldUploadResponse = 'description';
      fileUpload.completed = true;
      super.requestStateUpdate(true);
    }, 500);
  }

  /**
   * Retrieve File content as ArrayBuffer
   *
   * @param file
   * @returns
   */
  protected readFileContent(file: File): Promise<string | ArrayBuffer> {
    return new Promise<string | ArrayBuffer>((resolve, reject) => {
      const myReader: FileReader = new FileReader();

      myReader.onloadend = e => {
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
  protected async getFilesFromUploadArea(filesItems) {
    const folders = [];
    let entry: any;
    const collectFilesItems: File[] = [];

    for (let i = 0; i < filesItems.length; i++) {
      const uploadFileItem = filesItems[i];
      if (uploadFileItem.kind === 'file') {
        //Defensive code to validate if function exists in Browser
        //Collect all Folders into Array
        if (uploadFileItem.getAsEntry) {
          entry = uploadFileItem.getAsEntry();
          if (entry.isDirectory) {
            folders.push(entry);
          } else {
            const file = uploadFileItem.getAsFile();
            if (file) {
              file.fullPath = '';
              collectFilesItems.push(file);
            }
          }
        } else if (uploadFileItem.webkitGetAsEntry) {
          entry = uploadFileItem.webkitGetAsEntry();
          if (entry.isDirectory) {
            folders.push(entry);
          } else {
            const file = uploadFileItem.getAsFile();
            if (file) {
              file.fullPath = '';
              collectFilesItems.push(file);
            }
          }
        } else if ('function' == typeof uploadFileItem.getAsFile) {
          const file = uploadFileItem.getAsFile();
          if (file) {
            file.fullPath = '';
            collectFilesItems.push(file);
          }
        }
        continue;
      } else {
        const fileItem = uploadFileItem;
        if (fileItem) {
          fileItem.fullPath = '';
          collectFilesItems.push(fileItem);
        }
      }
    }
    //Collect Files from folder
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
  protected getFolderFiles(folders) {
    return new Promise<File[]>(resolve => {
      let reading = 0;
      const contents: File[] = [];
      folders.forEach(entry => {
        readEntry(entry, '');
      });

      function readEntry(entry, path) {
        if (entry.isDirectory) {
          readReaderContent(entry.createReader());
        } else {
          reading++;
          entry.file(file => {
            reading--;
            //Include Folder path where File is located
            file.fullPath = path;
            contents.push(file);

            if (reading === 0) {
              resolve(contents);
            }
          });
        }
      }

      function readReaderContent(reader) {
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
      }
    });
  }
}
