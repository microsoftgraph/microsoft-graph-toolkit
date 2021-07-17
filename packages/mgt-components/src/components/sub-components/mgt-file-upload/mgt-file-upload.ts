/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, property, css, TemplateResult } from 'lit-element';
import { styles } from './mgt-file-upload-css';
import { repeat } from 'lit-html/directives/repeat';
import { getSvg, SvgIcon } from '../../../utils/SvgHelper';
import { MgtBaseComponent } from '@microsoft/mgt-element';
import { ViewType } from '../../../graph/types';
import { DriveItem } from '@microsoft/microsoft-graph-types';

export { FluentDesignSystemProvider, FluentProgressRing } from '@fluentui/web-components';

/**
 * Configuration object for MgtFileList Properties.
 *
 * @export
 * @interface MgtFileListProperties
 */
export interface MgtFileListProperties {
  /**
   * Sets or gets whether the person component can use Contacts APIs to
   * find contacts and their images
   *
   * @type {string}
   */
  fileListQuery?: string;

  fileQueries?: string[];

  siteId?: string;

  driveId?: string;

  groupId?: string;

  itemId?: string;

  itemPath?: string;

  userId?: string;

  maxFileSize?: Number;

  enableFileUpload?: boolean;

  maxUploadFile?: Number;

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
      width: 100%;
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
      }

      .file-upload-button {
        background-color: rgb(243, 242, 241);
        width: 120px;
        height: 32px;
        text-align: center;
        display: table;
        float: right;
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
    `;
  }

  /**
   * allows developer to provide an array of files
   *
   * @type {File[]}
   * @memberof MgtFileUpload
   */
  @property({ type: Object })
  public files: File[];

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

  constructor() {
    super();
    this.files = [];
  }

  public render(): TemplateResult {
    console.log(this.fileListProperties);
    if (this.fileListProperties.enableFileUpload) {
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
            <span>${getSvg(SvgIcon.Upload)}Upload Files</span> 
          </div>
         </div>
         <div id="file-upload-Template">
            ${repeat(
              this.files,
              f => f.name,
              f => html`
                  ${this.renderFileTemplate(f)}
              `
            )}
         </div>
       `;
    }
  }

  private renderFileTemplate(fileUpload: File) {
    const view: ViewType = ViewType.image;
    const file: DriveItem = {
      name: fileUpload.name
    };
    return html`
      <div class='file-upload-table'>
        <div class='file-upload-cell'>
          <div class="file-upload-status">
            ${getSvg(SvgIcon.Fail)}
          </div>
          <mgt-file .fileDetails=${file} .view=${view}></mgt-file>   
        </div>
        <div class='file-upload-cell'>
          <div class='file-upload-table'>
            <div class='file-upload-cell'>
              ${fileUpload.name}
            </div>
          </div>
          <div class='file-upload-table'>
            <div class='file-upload-cell'>
              <div class='file-upload-table'>
                <div class='file-upload-cell'>
                  <div id="file-upload-bar">   
                    <div id="file-upload-progress"></div>
                  </div>
                </div>
                <div class='file-upload-cell' style="padding-left:5px">
                  <span>45%</span>
                </div>
                <div class='file-upload-cell'>
                  <span>${getSvg(SvgIcon.Cancel)}</span>
                </div>
              <div>
            </div>
          </div>
        </div>
      </div>
      `;
  }

  /**
   * Stop listeners from onDragOver event.
   * @param e
   */
  private handleonDragOver = e => {
    e.preventDefault();
    e.stopPropagation();
    if (e.dataTransfer.items && e.dataTransfer.items.length > 0) {
      //e.dataTransfer.dropEffect = e.dataTransfer.dropEffect = this.props.dropEffect;
    }
  };
  /**
   * Stop listeners from onDragEnter event, enable drag and drop view.
   * @param e
   */
  private handleonDragEnter = e => {
    e.preventDefault();
    e.stopPropagation();

    this._dragCounter++;
    if (e.dataTransfer.items && e.dataTransfer.items.length > 0) {
      //e.dataTransfer.dropEffect = this._dropEffect;
      const dragFileBorder = this.renderRoot.querySelector('#file-drag-border');
      dragFileBorder.classList.add('file-drag-border');
    }
  };

  /**
   * Stop listeners from ondragenter event, disable drag and drop view.
   * @param e
   */
  private handleonDragLeave = e => {
    e.preventDefault();
    e.stopPropagation();

    this._dragCounter--;
    if (this._dragCounter === 0) {
      const dragFileBorder = this.renderRoot.querySelector('#file-drag-border');
      dragFileBorder.classList.remove('file-drag-border');
    }
  };
  /**
   * Stop listeners from onDrop event and load files to property onDrop.
   * @param e
   */
  private handleonDrop = async e => {
    e.preventDefault();
    e.stopPropagation();

    const dragFileBorder = this.renderRoot.querySelector('#file-drag-border');
    dragFileBorder.classList.remove('file-drag-border');
    if (e.dataTransfer && e.dataTransfer.items) {
      this.onDrop(await this.getFiles(e));
    }
    e.dataTransfer.clearData();
    this._dragCounter = 0;
  };

  protected onDrop(files) {
    if (files.length > 0) {
      this.files = files;
      for (
        var i = 0;
        i < (this.fileListProperties.maxFileSize > files.length ? files.length : this.fileListProperties.maxFileSize);
        i++
      ) {
        console.log('Filename: ' + files[i].name);
        console.log('Path: ' + files[i].fullPath);
      }
    }
  }

  /**
   *  Get files objects and includes fullPath properties in File Object
   * @param e
   */
  protected async getFiles(e) {
    const uploadFilesItems = e.dataTransfer.items;
    const directory = [];
    let entry: any;
    const files: File[] = [];
    for (let i = 0; i < uploadFilesItems.length; i++) {
      const item = uploadFilesItems[i];
      if (item.kind === 'file') {
        if (item.getAsEntry) {
          entry = item.getAsEntry();
          if (entry.isDirectory) {
            directory.push(entry);
          } else {
            const file = item.getAsFile();
            if (file) {
              file.fullPath = '';
              files.push(file);
            }
          }
        } else if (item.webkitGetAsEntry) {
          entry = item.webkitGetAsEntry();
          if (entry.isDirectory) {
            directory.push(entry);
          } else {
            const file = item.getAsFile();
            if (file) {
              file.fullPath = '';
              files.push(file);
            }
          }
        } else if ('function' == typeof item.getAsFile) {
          const file = item.getAsFile();
          if (file) {
            file.fullPath = '';
            files.push(file);
          }
        }
        continue;
      }
    }
    if (directory.length > 0) {
      const entryContent = await this.readEntryContentAsync(directory);
      files.push(...entryContent);
    }
    return files;
  }

  private readEntryContentAsync = Directory => {
    return new Promise<File[]>((resolve, reject) => {
      let reading = 0;
      const contents: File[] = [];
      Directory.forEach(entry => {
        readEntry(entry, '');
      });

      function readEntry(entry, path) {
        if (entry.isDirectory) {
          readReaderContent(entry.createReader());
        } else {
          reading++;
          entry.file(file => {
            reading--;
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
  };
}
