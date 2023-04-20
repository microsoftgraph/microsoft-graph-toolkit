/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { fluentProgress } from '@fluentui/web-components';
import { html, nothing } from 'lit';
import { property } from 'lit/decorators.js';
import { customElement, MgtTemplatedComponent, mgtHtml } from '@microsoft/mgt-element';
import { MgtFileUploadItem } from './mgt-file-upload';
import { ViewType } from '../../../graph/types';
import { registerFluentComponents } from '../../../utils/FluentComponents';
import { classMap } from 'lit/directives/class-map.js';
import { getSvg, SvgIcon } from '../../../utils/SvgHelper';
import { styles } from './mgt-file-upload-progress-css';
import { strings } from './strings';

registerFluentComponents(fluentProgress);

/**
 * Component used to render progress of file upload
 *
 * @fires {CustomEvent<MgtFileUploadItem>} clearnotification - Fired when notification is cleared
 *
 * @export
 * @class mgt-component
 * @extends {MgtTemplatedComponent}
 */
@customElement('file-upload-progress')
export class MgtFileUploadProgress extends MgtTemplatedComponent {
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
   * Array of progress items to render
   *
   * @type {MgtFileUploadItem[]}
   * @memberof MgtComponent
   */
  @property({
    attribute: null
  })
  public progressItems: MgtFileUploadItem[] = [];

  /**
   * Synchronizes property values when attributes change.
   *
   * @param {string} name
   * @param {string} oldValue
   * @param {string} newValue
   * @memberof MgtPersonCard
   */
  public attributeChangedCallback(name: string, oldValue: string, newValue: string) {
    super.attributeChangedCallback(name, oldValue, newValue);

    // TODO: handle when an attribute changes.
    //
    // Ex: load data when the name attribute changes
    // if (name === 'person-id' && oldval !== newval){
    //  this.loadData();
    // }
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    return html`
      <div class='root'>
        <div class="file-upload-Template">
          ${this.renderFolderTemplate(this.progressItems)}
        </div>
      </div>
    `;
  }

  /**
   * Render Folder structure of files to upload
   *
   * @param fileItems
   * @returns
   */
  protected renderFolderTemplate(fileItems: MgtFileUploadItem[]) {
    const folderStructure: string[] = [];
    if (fileItems.length > 0) {
      const templateFileItems = fileItems.map(fileItem => {
        if (folderStructure.indexOf(fileItem.fullPath.substring(0, fileItem.fullPath.lastIndexOf('/'))) === -1) {
          if (fileItem.fullPath.substring(0, fileItem.fullPath.lastIndexOf('/')) !== '') {
            folderStructure.push(fileItem.fullPath.substring(0, fileItem.fullPath.lastIndexOf('/')));
            return mgtHtml`
            <div class='file-upload-table'>
              <div class='file-upload-cell'>
                <mgt-file
                  .fileDetails=${{
                    name: fileItem.fullPath.substring(1, fileItem.fullPath.lastIndexOf('/')),
                    folder: 'Folder'
                  }}
                  .view=${ViewType.oneline}
                  class="mgt-file-item">
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
      return html`${templateFileItems}`;
    }
    return nothing;
  }

  /**
   * Render file upload area
   *
   * @param fileItem
   * @returns
   */
  protected renderFileTemplate(fileItem: MgtFileUploadItem, folderTabStyle: string) {
    const completed = classMap({
      'file-upload-table': true,
      upload: fileItem.completed
    });
    const folder =
      folderTabStyle + (fileItem.fieldUploadResponse === 'lastModifiedDateTime' ? ' file-upload-dialog-success' : '');

    const description = classMap({
      description: fileItem.fieldUploadResponse === 'description'
    });

    const completedTemplate = !fileItem.completed
      ? this.renderFileUploadTemplate(fileItem)
      : this.renderCompletedUploadButton(fileItem);

    return mgtHtml`
        <div class="${completed}">
          <div class="${folder}">
            <div class='file-upload-cell'>
              <div class="${description}">
                <div class="file-upload-status">
                  ${fileItem.iconStatus}
                </div>
                <mgt-file
                  .fileDetails=${fileItem.driveItem}
                  .view=${fileItem.view}
                  .line2Property=${fileItem.fieldUploadResponse}
                  part="upload"
                  class="mgt-file-item">
                </mgt-file>
              </div>
            </div>
              ${completedTemplate}
            </div>
          </div>
        </div>`;
  }

  /**
   * Render file upload progress
   *
   * @param fileItem
   * @returns
   */
  protected renderFileUploadTemplate(fileItem: MgtFileUploadItem) {
    const completed = classMap({
      'file-upload-table': true,
      upload: fileItem.completed
    });
    return html`
    <div class='file-upload-cell'>
      <div class='file-upload-table file-upload-name' >
        <div class='file-upload-cell'>
          <div
            title="${fileItem.file.name}"
            class='file-upload-filename'>
            ${fileItem.file.name}
          </div>
        </div>
      </div>
      <div class='file-upload-table'>
        <div class='file-upload-cell'>
          <div class="${completed}">
            <fluent-progress class="file-upload-bar" value="${fileItem.percent}"></fluent-progress>
            <div class='file-upload-cell percent-indicator'>
              <span>${fileItem.percent}%</span>
              <span
                class="file-upload-cancel"
                @click=${() => this.deleteFileUploadSession(fileItem)}>
                ${getSvg(SvgIcon.Cancel)}
              </span>
            </div>
          <div>
        </div>
      </div>
    </div>
    `;
  }

  private renderCompletedUploadButton(fileItem: MgtFileUploadItem) {
    const ariaLabel = `${strings.clearNotification} ${fileItem.file.name}`;
    return html`
      <span class="file-upload-cell">
        <fluent-button
          part="file-upload-remove"
          appearance="stealth"
          aria-label=${ariaLabel}
          @click=${() => this.removeNotification(fileItem)}
        >
          ${getSvg(SvgIcon.Cancel)}
        </fluent-button>
      </span>
    `;
  }

  private removeNotification(fileItem: MgtFileUploadItem) {
    this.fireCustomEvent('clearnotification', fileItem);
  }

  /**
   * Function to delete existing file upload sessions
   *
   * @param fileItem
   */
  protected deleteFileUploadSession(fileItem: MgtFileUploadItem) {
    try {
      if (fileItem.uploadUrl !== undefined) {
        // Responses that confirm cancelation of session.
        // 404 means (The upload session was not found/The resource could not be found/)
        // 409 means The resource has changed since the caller last read it; usually an eTag mismatch
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
   * Change the state of Mgt-File icon upload to Fail
   *
   * @param fileUpload
   */
  protected setUploadFail(fileUpload: MgtFileUploadItem, errorMessage: string) {
    fileUpload.iconStatus = getSvg(SvgIcon.Fail);
    fileUpload.view = ViewType.twolines;
    fileUpload.driveItem.description = errorMessage;
    fileUpload.fieldUploadResponse = 'description';
    fileUpload.completed = true;
    void super.requestStateUpdate(true);
  }
}
