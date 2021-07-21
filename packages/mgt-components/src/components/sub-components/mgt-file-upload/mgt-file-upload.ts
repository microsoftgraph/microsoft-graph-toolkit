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
import { IGraph, MgtBaseComponent, prepScopes } from '@microsoft/mgt-element';
import { ViewType } from '../../../graph/types';
import { DriveItem } from '@microsoft/microsoft-graph-types';

export { FluentDesignSystemProvider, FluentProgressRing } from '@fluentui/web-components';

/**
 * MgtFileUpload upload item lifecycle object.
 *
 * @export
 * @interface MgtFileUploadItem
 */
export interface MgtFileUploadItem {
  /**
   * Graph Url to upload file
   *
   */
  GraphUrl?: string;

  /**
   * Session url to keep upload progress open untill all chuncks are sent
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
 * MgtFileUpload configuration object with MgtFileList Properties.
 *
 * @export
 * @interface MgtFileUploadConfig
 */
export interface MgtFileUploadConfig {
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
 * A component to upload files to OneDrive or SharePoint Sites
 *
 * @export
 * @class MgtFileUpload
 * @extends {MgtBaseComponent}
 *
 * @cssprop --file-upload-border- {String} File upload border top style
 * @cssprop --file-upload-background-color - {Color} File upload background color with opacity style
 * @cssprop --file-upload-button-text-align - {text-align} Upload button aligment using -webkit-[position]
 * @cssprop --file-upload-button-background-color - {Color} Background color of upload button
 * @cssprop --file-upload-button-color - {Color} Text color of upload button
 * @cssprop --file-upload-progress-background-color - {Color} progress background color
 * @cssprop --file-upload-progressBar-background-color - {Color} progressBar background color
 * @cssprop --file-upload-item-background-color - {Color} Background color of upload file
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

  private _dragCounter: number = 0;
  private _dropEffect: string = 'copy';

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
              accept=""
              @change="${this.onFileUploadChange}"
            />
            <span @click=${this.onFileUploadClick}>${getSvg(SvgIcon.Upload)}${strings.buttonUploadFile}</span> 
          </div>
         </div>
         <div id="file-upload-Template">
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
        <div class='file-upload-table'>
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
   * Handle the click event to cancel upload file
   *
   * @param fileItem
   */
  protected onFileUploadCancelClick(fileItem: MgtFileUploadItem) {
    if (fileItem.uploadUrl !== undefined) {
      this.fileUploadList.graph.client.api(fileItem.uploadUrl).delete(async response => {
        if (response === null) {
          fileItem.uploadUrl = undefined;
        }
      });
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
   * Stop listeners from onDrop event and load files to property onDrop.
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
   * Initialize MgtFileUploadItem with files to upload
   *
   * @param inputFiles
   */
  protected async getSelectedFiles(files: File[]) {
    const maxFiles =
      this.fileUploadList.maxUploadFile > files.length ? files.length : this.fileUploadList.maxUploadFile;
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
    //Send Files to upload asyncronous (Multiple Uploads)
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
  protected async sendFileItemGraph(fileItem: MgtFileUploadItem) {
    if (fileItem.file.size < 4 * 1024 * 1024) {
      const scopes = 'files.readwrite';
      try {
        const driveItem: DriveItem = await this.fileUploadList.graph
          .api(`me/drive/root:/${fileItem.file.name}:/content`)
          .middlewareOptions(prepScopes(scopes))
          .put(fileItem.file);
        fileItem.driveItem = driveItem;
        this.setUploadSuccess(fileItem);
      } catch (error) {
        this.setUploadFail(fileItem, strings.failUploadFile);
      }
    } else {
      fileItem.GraphUrl = `me/drive/root:/${fileItem.file.name}:/createUploadSession`;
      const driveItem: DriveItem = await this.sendLargeFileGraph(this.fileUploadList.graph, fileItem);
      if (driveItem !== undefined) {
        fileItem.driveItem = driveItem;
        this.setUploadSuccess(fileItem);
      }
    }
  }

  /**
   * Initialize Upload file using Session url "uploadUrl"
   *
   * @param Graph
   * @param fileItem
   * @returns
   */
  protected async sendLargeFileGraph(Graph: IGraph, fileItem: MgtFileUploadItem) {
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
  protected async sendSessionUrlGraph(Graph: IGraph, fileItem: MgtFileUploadItem) {
    let minSize = 0;
    let maxSize = 5 * 327680; //define max chunk size to upload

    while (fileItem.file.size > minSize) {
      if (fileItem.mimeStreamString === undefined) {
        fileItem.mimeStreamString = (await this.readFileContent(fileItem.file)) as ArrayBuffer;
      }
      //Graph client API uses Buffer package to manage ArrayBuffer, change to Blob avoids external package dependency
      const fileSlice: Blob = new Blob([fileItem.mimeStreamString.slice(minSize, maxSize)]);
      fileItem.percent = Math.round((maxSize / fileItem.file.size) * 100);
      super.requestStateUpdate(true);

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
  protected async sendFileChuncksGraph(
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
      //If upload chunck Graph call fails delete current Session
      if (fileItem.uploadUrl !== undefined) {
        this.fileUploadList.graph.client.api(fileItem.uploadUrl).delete(async response => {
          if (response === null) {
            fileItem.uploadUrl = undefined;
            this.setUploadFail(fileItem, strings.failUploadFile);
          }
        });
      } else {
        this.setUploadFail(fileItem, strings.cancelUploadFile);
      }
      return undefined;
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
  protected setUploadFail(fileUpload, errorMessage: string) {
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
    const maxFiles =
      this.fileUploadList.maxUploadFile > filesItems.length ? filesItems.length : this.fileUploadList.maxUploadFile;

    for (let i = 0; i < maxFiles; i++) {
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
