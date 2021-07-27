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
import { IGraph, MgtBaseComponent } from '@microsoft/mgt-element';
import { ViewType } from '../../../graph/types';
import { DriveItem } from '@microsoft/microsoft-graph-types';
import { getUploadSession, sendFileContent, sendFileChunck, deleteSessionFile } from '../../../graph/graph.files';

export { FluentDesignSystemProvider, FluentProgressRing } from '@fluentui/web-components';

/**
 * MgtFileUpload upload item lifecycle object.
 *
 * @export
 * @interface MgtFileUploadItem
 */
export interface MgtFileUploadItem {
  /**
   * MS Graph Provider
   */
  GraphUrl?: string;

  /**
   * Session url to keep upload progress open untill all chuncks are sent
   */
  uploadUrl?: string;

  /**
   *  Upload file progress value
   */
  percent?: number;

  /**
   *  Output "Success" or "Fail" icon base on upload response
   */
  iconStatus?: TemplateResult;

  /**
   * File object to be upload.
   */
  file?: File;

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
 * @cssprop --file-upload-button-color - {Color} Text color of upload button
 * @cssprop --file-upload-progress-background-color - {Color} progress background color
 * @cssprop --file-upload-progressbar-background-color - {Color} progressBar background color
 * @cssprop --file-item-margin - {String} File item margin
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
         <div id="file-upload-border" >
            
         </div>
         <div class="file-upload-button">
          <div>
            <input
              id="file-upload-input"
              type="file"
              multiple="true" 
              @change="${this.onFileUploadChange}"
            />
            <span @click=${this.onFileUploadClick}>${getSvg(SvgIcon.Upload)}${strings.buttonUploadFile}</span> 
          </div>
         </div>
         <div class="file-upload-Template">
         ${this.renderFileTemplate(this.filesToUpload)}
         </div>
       `;
  }

  /**
   * Render file upload area
   *
   * @param fileItems
   * @returns
   */
  protected renderFileTemplate(fileItems: MgtFileUploadItem[]) {
    if (fileItems.length > 0) {
      const TemplateFileItems = fileItems.map(fileItem => {
        return html`
        <div class='file-upload-table' style="${fileItem.completed ? 'width: 100%;' : null}">
          <div class='file-upload-cell'>
            <div style=${fileItem.fieldUploadResponse === 'description' ? 'opacity: 0.5;' : null}>
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
        `;
      });
      return html`${TemplateFileItems}`;
    } else {
      return null;
    }
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
      <div class='file-upload-table'>
        <div class='file-upload-cell'>
          ${fileItem.file.name}
        </div>
      </div>
      <div class='file-upload-table'>
        <div class='file-upload-cell'>
          <div class='file-upload-table'>
            <div class='file-upload-cell'>
              <div id="file-upload-bar">   
                <div id="file-upload-progress" style="width: ${fileItem.percent}%;"></div>
              </div>
            </div>
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
      dragFileBorder.style.display = 'inline-block';
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
      dragFileBorder.style.display = 'none';
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
    dragFileBorder.style.display = 'none';
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

    //Collect ongoing upload files
    this.filesToUpload.forEach(async fileItem => {
      if (!fileItem.completed) {
        fileItems.push(fileItem);
      }
    });

    for (var i = 0; i < files.length; i++) {
      if (fileItems.filter(item => item.file.name === files[i].name).length === 0) {
        //Initialize variable for File validation
        let acceptFile = true;

        //Exclude file based on max files upload allowed
        if (fileItems.length >= this.fileUploadList.maxUploadFile) {
          acceptFile = false;
        }

        //Exclude file based on max file size allowed
        if (this.fileUploadList.maxFileSize !== undefined && acceptFile) {
          if (files[i].size > this.fileUploadList.maxFileSize * 1024) {
            acceptFile = false;
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
            }
          }
        }

        //Collect accepted files
        if (acceptFile) {
          //Initialize MgtFileUploadItem Life cycle
          fileItems.push({
            file: files[i],
            driveItem: {
              name: files[i].name
            },
            iconStatus: null,
            percent: 1,
            view: ViewType.image,
            completed: false,
            maxSize: this._maxChunckSize,
            minSize: 0
          });
        }
      }
    }
    this.filesToUpload = fileItems;
    // Send multiple Files to upload
    this.filesToUpload.forEach(async fileItem => {
      await this.sendFileItemGraph(fileItem);
    });
  }

  /**
   * Get GraphQuery based on pre defined parameters.
   *
   * @param fileItem
   * @returns
   */
  protected getGrapQuery(fileItem: any) {
    let fullPath = fileItem.fullPath === '' ? '/' + fileItem.name : fileItem.fullPath;
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
    if (fileItem.file.size < this._maxChunckSize) {
      try {
        if (!fileItem.completed) {
          fileItem.driveItem = await sendFileContent(
            graph,
            `${this.getGrapQuery(fileItem.file)}:/content`,
            fileItem.file
          );
          if (fileItem.driveItem !== null) {
            this.setUploadSuccess(fileItem);
          } else {
            this.setUploadFail(fileItem, strings.failUploadFile);
          }
        }
      } catch (error) {
        this.setUploadFail(fileItem, strings.failUploadFile);
      }
    } else {
      fileItem.GraphUrl = `${this.getGrapQuery(fileItem.file)}:/createUploadSession`;
      if (fileItem.uploadUrl === undefined) {
        const response = await getUploadSession(graph, fileItem.GraphUrl);
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
          //
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
      fileUpload.iconStatus = getSvg(SvgIcon.Sucess);
      fileUpload.view = ViewType.twolines;
      fileUpload.fieldUploadResponse = 'lastModifiedDateTime';
      fileUpload.completed = true;
      super.requestStateUpdate(true);
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
