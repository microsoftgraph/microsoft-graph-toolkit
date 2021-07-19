/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, property, css, TemplateResult } from 'lit-element';
import { styles } from './mgt-file-upload-css';
import { getSvg, SvgIcon } from '../../../utils/SvgHelper';
import { IGraph, MgtBaseComponent, prepScopes } from '@microsoft/mgt-element';
import { ViewType } from '../../../graph/types';
import { DriveItem } from '@microsoft/microsoft-graph-types';

export { FluentDesignSystemProvider, FluentProgressRing } from '@fluentui/web-components';

export interface MgtFileUploadItem {
  /**
   * Graph Url to upload file
   *
   */
  GraphUrl?: string;

  /**
   * Upload url keeps session open untill all chuncks are sent
   *
   */
  uploadUrl?: string;

  /**
   *  Upload file progress value
   *
   */
  percent?: number;

  /**
   *  Output "Success" or "Fail" icon base on upload response
   *
   */
  iconStatus?: TemplateResult;

  /**
   * File object to be upload.
   *
   */
  file?: File;

  /**
   * Mgt-File View state change on upload response
   *
   */
  view?: ViewType;

  /**
   * Manipulate fileDetails on upload lifecycle
   *
   */
  driveItem?: DriveItem;

  /**
   * Mgt-File line2Property output field message
   *
   */
  fieldUploadResponse?: string;

  /**
   * Validates state of upload progress
   *
   */
  completed?: boolean;

  /**
   * Load large Files into ArrayBuffer to send by chuncks
   *
   */
  mimeStreamString?: ArrayBuffer;
}

/**
 * Configuration object for MgtFileList Properties.
 *
 * @export
 * @interface MgtFileListProperties
 */
export interface MgtFileListProperties {
  graph: IGraph;

  /**
   * Sets or gets whether the person component can use Contacts APIs to
   * find contacts and their images
   *
   * @type {string}
   */
  fileListQuery?: string;

  /**
   *
   *
   *  @type {string[]}
   */
  fileQueries?: string[];

  /**
   *
   *
   *  @type {string}
   */
  siteId?: string;

  /**
   *
   *
   *  @type {string}
   */
  driveId?: string;

  /**
   *
   *
   *  @type {string}
   */
  groupId?: string;

  /**
   *
   *
   *  @type {string}
   */
  itemId?: string;

  /**
   *
   *
   *  @type {string}
   */
  itemPath?: string;

  /**
   *
   *
   *  @type {string}
   */
  userId?: string;

  /**
   *
   *
   *  @type {Number}
   */
  maxFileSize?: Number;

  /**
   *
   *
   *  @type {boolean}
   */
  enableFileUpload?: boolean;

  /**
   *
   *
   *  @type {Number}
   */
  maxUploadFile?: Number;

  /**
   *
   *
   *  @type {string[]}
   */
  excludedFileExtensions?: string[];
}

/**
 * A component to create flyout anchored to an element
 *
 * @export
 * @class MgtFileUpload
 * @extends {LitElement}
 */
@customElement('mgt-file-upload')
export class MgtFileUpload extends MgtBaseComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return css`
    .file-upload-input {
      width: 0.1px;
      height: 0.1px;
      opacity: 0;
      overflow: hidden;
      position: absolute;
      z-index: -1;
    }
    
    .file-upload-table {
      display:table;
      width:260px;
    }
    .file-upload-cell {
        padding:1px 0px 1px 1px;
        display:table-cell;
        width:50%;
        vertical-align:middle;
        position:relative;
    }

    .file-upload-status{
      position: absolute;
      top: 12px;
      left: 16px;
    }

    #file-upload-bar {
      max-width: 200px;
      background-color: rgb(243, 242, 241);
      border-radius:12px;
    }
    
    #file-upload-progress {
      height: 5px;
      background-color: #0078d4;
      border-radius:12px;
    }
    
    .file-drag-border {
        border: dashed #0078d4 1px;
        background-color: rgba(0, 120, 212, 0.1);
        position: absolute;
        top: 0;
        bottom: 0;
        left: 0;
        right: 0;
        z-index: 1;
      }
    
      .file-upload-border{
        width: 100%;
        text-align: -webkit-right;
      }

      .file-upload-button {
        background-color: rgb(243, 242, 241);
        width: 120px;
        height: 32px;
        text-align: center;
        display: table;
        margin-top: 39px;
        margin-right: 16px;
        cursor: pointer;
        font-size: 14px;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      }
      .file-upload-button span {
        vertical-align: middle;
        display: table-cell;
      }

      .file-upload-cancel{
        cursor: pointer;
      }
    `;
  }

  /**
   * Allows developer to provide an array of files to upload
   *
   * @type {MgtFileUploadItem[]}
   * @memberof MgtFileUpload
   */
  @property({ type: Object })
  public filesToUpload: MgtFileUploadItem[];

  /**
   * List of mgt-File-List Properties used for upload of files.
   * @type {MgtFileListProperties}
   * @memberof MgtFileUpload
   */
  @property({ type: Object })
  public fileListProperties: MgtFileListProperties;

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

  private _dragCounter: number = 0;
  private _dropEffect = 'copy';
  private _uploadStatus: boolean = false;

  constructor() {
    super();
    this.filesToUpload = [];
  }

  /**
   * Override requestStateUpdate to include clearstate.
   *
   * @memberof MgtFileUpload
   */
  protected requestStateUpdate(force?: boolean) {
    //this.clearState();
    return super.requestStateUpdate(force);
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
         <div id="file-drag-border" ></div>
         <div class="file-upload-border">
          <div class="file-upload-button" >
          <input
                class="file-upload-input"
                id="file-upload-input"
                type="file"
                multiple="true" 
                accept=""
                @change="${this.onFileUploadChange}"
              />
            <span @click=${this.onFileUploadClick}>${getSvg(SvgIcon.Upload)}Upload Files</span> 
          </div>
         </div>
         <div id="file-upload-Template">
         ${this.renderFileTemplate(this.filesToUpload)}
         </div>
       `;
  }

  /**
   * Render File Upload Area
   *
   * @param fileUpload
   * @returns
   */
  private renderFileTemplate(fileItems: MgtFileUploadItem[]) {
    if (fileItems.length > 0) {
      const TemplateFileItems = fileItems.map(fileItem => {
        return html`
        <div class='file-upload-table'>
          <div class='file-upload-cell'>
            <div style=${fileItem.fieldUploadResponse === 'description' ? 'opacity: 0.5;' : null}>
              <div class="file-upload-status">
                ${fileItem.iconStatus}
              </div>
              <mgt-file 
                .fileDetails=${fileItem.driveItem} 
                .view=${fileItem.view} 
                .line2Property=${fileItem.fieldUploadResponse}>
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
   * Render File Upload Progress
   *
   * @param fileItem
   * @returns
   */
  private renderFileUploadTemplate(fileItem: MgtFileUploadItem) {
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
            </div>
            <div class='file-upload-cell' >
              <span 
                class="file-upload-cancel" 
                @click=${e => this.onFileUploadCancelClick(fileItem)}>
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
   *
   * @param e
   * @returns
   */
  private async onFileUploadChange(e) {
    if (!e || e.target.files.length < 1) {
      return;
    } else {
      this.getSelectedFiles(await this.getFilesFromUploadArea(e.target.files));
    }
  }

  /**
   *
   */
  private onFileUploadClick() {
    const uploadInput: HTMLElement = this.renderRoot.querySelector('#file-upload-input');
    uploadInput.click();
  }

  /**
   * Cancel upload of file
   *
   * @param uploadFile
   */
  public onFileUploadCancelClick(fileItem: MgtFileUploadItem) {
    this.fileListProperties.graph.client.api(fileItem.uploadUrl).delete(async response => {
      if (response === null) {
        fileItem.uploadUrl = undefined;
      }
    });
  }

  /**
   * Stop listeners from onDragOver event.
   *
   * @param e
   */
  protected handleonDragOver = e => {
    e.preventDefault();
    e.stopPropagation();

    if (e.dataTransfer.items && e.dataTransfer.items.length > 0) {
      e.dataTransfer.dropEffect = e.dataTransfer.dropEffect = this._dropEffect;
    }
  };

  /**
   * Stop listeners from onDragEnter event, enable drag and drop view.
   *
   * @param e
   */
  protected handleonDragEnter = e => {
    e.preventDefault();
    e.stopPropagation();

    this._dragCounter++;
    if (e.dataTransfer.items && e.dataTransfer.items.length > 0) {
      e.dataTransfer.dropEffect = this._dropEffect;
      const dragFileBorder: HTMLElement = this.renderRoot.querySelector('#file-drag-border');
      dragFileBorder.classList.add('file-drag-border');
    }
  };

  /**
   * Stop listeners from ondragenter event, disable drag and drop view.
   *
   * @param e
   */
  protected handleonDragLeave = e => {
    e.preventDefault();
    e.stopPropagation();

    this._dragCounter--;
    if (this._dragCounter === 0) {
      const dragFileBorder: HTMLElement = this.renderRoot.querySelector('#file-drag-border');
      dragFileBorder.classList.remove('file-drag-border');
    }
  };
  /**
   * Stop listeners from onDrop event and load files to property onDrop.
   *
   * @param e
   */
  protected handleonDrop = async e => {
    e.preventDefault();
    e.stopPropagation();

    const dragFileBorder = this.renderRoot.querySelector('#file-drag-border');
    dragFileBorder.classList.remove('file-drag-border');
    if (e.dataTransfer && e.dataTransfer.items) {
      this.getSelectedFiles(await this.getFilesFromUploadArea(e.dataTransfer.items));
    }
    e.dataTransfer.clearData();
    this._dragCounter = 0;
  };

  /**
   * Initialize MgtFileUploadItem with files to upload
   *
   * @param inputFiles
   */
  protected async getSelectedFiles(files: File[]) {
    const maxFiles =
      this.fileListProperties.maxUploadFile > files.length ? files.length : this.fileListProperties.maxUploadFile;
    let fileItems: MgtFileUploadItem[] = [];

    if (files.length > 0) {
      for (var i = 0; i < maxFiles; i++) {
        //Initialize MgtFileUploadItem with files
        let fileItem: MgtFileUploadItem = {
          file: files[i],
          driveItem: {
            name: files[i].name
          },
          iconStatus: null,
          percent: 1,
          view: ViewType.image,
          completed: false
        };
        fileItems.push(fileItem);
      }
    }
    this.filesToUpload = fileItems;
    this.filesToUpload.forEach(async fileItem => {
      await this.sendFileItemGraph(fileItem);
    });
  }

  /**
   * Select Upload Graph Method based on file length
   *
   * @param fileUpload
   * @returns
   */
  private async sendFileItemGraph(fileItem: MgtFileUploadItem) {
    if (fileItem.file.size < 4 * 1024 * 1024) {
      const scopes = 'files.readwrite';
      try {
        const driveItem: DriveItem = await this.fileListProperties.graph
          .api(`me/drive/root:/${fileItem.file.name}:/content`)
          .middlewareOptions(prepScopes(scopes))
          .put(fileItem.file);
        fileItem.driveItem = driveItem;
        this.setUploadSuccess(fileItem);
      } catch (error) {
        this.setUploadFail(fileItem, 'File upload fail.');
      }
    } else {
      fileItem.GraphUrl = `me/drive/root:/${fileItem.file.name}:/createUploadSession`;
      const driveItem: DriveItem = await this.sendLargeFileGraph(this.fileListProperties.graph, fileItem);
      if (driveItem !== undefined) {
        fileItem.driveItem = driveItem;
        this.setUploadSuccess(fileItem);
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
    this.requestStateUpdate(true);
    setTimeout(() => {
      fileUpload.iconStatus = getSvg(SvgIcon.Sucess);
      fileUpload.view = ViewType.twolines;
      fileUpload.fieldUploadResponse = 'lastModifiedDateTime';
      fileUpload.completed = true;
      this.requestStateUpdate(true);
    }, 500);
  }

  /**
   * Change the state of Mgt-File icon upload to Fail
   *
   * @param fileUpload
   */
  private setUploadFail(fileUpload, errorMessage: string) {
    setTimeout(() => {
      fileUpload.iconStatus = getSvg(SvgIcon.Fail);
      fileUpload.view = ViewType.twolines;
      fileUpload.driveItem.description = errorMessage;
      fileUpload.fieldUploadResponse = 'description';
      fileUpload.completed = true;
      this.requestStateUpdate(true);
    }, 500);
  }

  /**
   * Initialize Upload file using Session url "uploadUrl"
   *
   * @param Graph
   * @param fileItem
   * @returns
   */
  public async sendLargeFileGraph(Graph: IGraph, fileItem: MgtFileUploadItem) {
    const graph = Graph;
    const sessionOptions = {
      item: {
        '@microsoft.graph.conflictBehavior': 'rename'
      }
    };
    const scopes = 'files.readwrite';
    try {
      return graph
        .api('/' + fileItem.GraphUrl)
        .middlewareOptions(prepScopes(scopes))
        .post(JSON.stringify(sessionOptions))
        .then(
          async (response): Promise<DriveItem> => {
            try {
              //uploadUrl keeps upload Session open untill all chuncks are sent
              fileItem.uploadUrl = response.uploadUrl;
              const driveItem = await this.sendSessionUrlGraph(Graph, fileItem);
              return driveItem;
            } catch {
              return undefined;
            }
          }
        );
    } catch {
      return undefined;
    }
  }

  /**
   * Manage slices of File to upload file by chuncks using Graph and Session Url
   *
   * @param Graph
   * @param fileItem
   * @returns
   */
  private async sendSessionUrlGraph(Graph: IGraph, fileItem: MgtFileUploadItem) {
    let minSize = 0;
    let maxSize = 5 * 327680; //define max chunk size to upload

    while (fileItem.file.size > minSize) {
      if (fileItem.mimeStreamString === undefined) {
        fileItem.mimeStreamString = (await this.readFileContent(fileItem.file)) as ArrayBuffer;
      }
      const fileSlice: Blob = new Blob([fileItem.mimeStreamString.slice(minSize, maxSize)]);
      fileItem.percent = Math.round((maxSize / fileItem.file.size) * 100);
      this.requestStateUpdate(true);

      if (fileItem.uploadUrl !== undefined) {
        const driveItem = await this.sendFileChuncksGraph(
          Graph,
          fileItem,
          `${maxSize - minSize}`,
          `bytes ${minSize}-${maxSize - 1}/${fileItem.file.size}`,
          fileSlice
        );
        if (driveItem === undefined) {
          return undefined;
        } else if (driveItem.id !== undefined) {
          return driveItem;
        }
      } else {
        return undefined;
      }

      minSize = maxSize;
      maxSize += 5 * 327680;
      if (maxSize > fileItem.file.size) {
        maxSize = fileItem.file.size;
      }
    }
  }

  /**
   * Upload File chunck by Graph using Session Url
   *
   * @param Graph
   * @param uploadUrl
   * @param contentLength
   * @param contentRange
   * @param file
   * @returns
   */
  private async sendFileChuncksGraph(
    Graph: IGraph,
    fileItem: MgtFileUploadItem,
    contentLength: string,
    contentRange: string,
    file: Blob
  ) {
    const graph = Graph;
    const scopes = 'files.readwrite';
    const header = {
      'Content-Length': contentLength,
      'Content-Range': contentRange
    };
    try {
      const response = await graph.client
        .api(fileItem.uploadUrl)
        .middlewareOptions(prepScopes(scopes))
        .headers(header)
        .put(file);
      return response;
    } catch {
      //If upload chunch Graph call fails delete current Session
      if (fileItem.uploadUrl !== undefined) {
        this.fileListProperties.graph.client.api(fileItem.uploadUrl).delete(async response => {
          if (response === null) {
            fileItem.uploadUrl = undefined;
            this.setUploadFail(fileItem, 'File upload fail.');
          }
        });
      } else {
        this.setUploadFail(fileItem, 'File cancel.');
      }
      return undefined;
    }
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
    // Collect Folders Array
    const folders = [];
    let entry: any;
    const collectFilesItems: File[] = [];
    const maxFiles =
      this.fileListProperties.maxUploadFile > filesItems.length
        ? filesItems.length
        : this.fileListProperties.maxUploadFile;

    for (let i = 0; i < maxFiles; i++) {
      const uploadFileItem = filesItems[i];
      if (uploadFileItem.kind === 'file') {
        //Defensive code to validate if function exists in Browser
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
    if (folders.length > 0) {
      const folderFiles = await this.getFolderFiles(folders);
      collectFilesItems.push(...folderFiles);
    }
    return collectFilesItems;
  }

  /**
   * Retrieve Files from Folder and Subfolders to array.
   *
   * @param folders
   * @returns
   */
  private getFolderFiles(folders) {
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
