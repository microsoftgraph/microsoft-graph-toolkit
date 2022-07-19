/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { SharedInsight } from '@microsoft/microsoft-graph-types';
import { customElement, html, TemplateResult } from 'lit-element';
import { BasePersonCardSection } from '../BasePersonCardSection';
import { getFileTypeIconUri } from '../../../../styles/fluent-icons';
import { getSvg, SvgIcon } from '../../../../utils/SvgHelper';
import { getRelativeDisplayDate } from '../../../../utils/Utils';
import { styles } from './mgt-person-card-files-css';
import { strings } from './strings';

/**
 * The files subsection of the person card
 *
 * @export
 * @class MgtPersonCardProfile
 * @extends {MgtTemplatedComponent}
 */
@customElement('mgt-person-card-files')
export class MgtPersonCardFiles extends BasePersonCardSection {
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

  private _files: SharedInsight[];

  public constructor(files: SharedInsight[]) {
    super();
    this._files = files;
  }

  /**
   * The name for display in the overview section.
   *
   * @readonly
   * @type {string}
   * @memberof MgtPersonCardFiles
   */
  public get displayName(): string {
    return this.strings.filesSectionTitle;
  }

  /**
   * Reset any state in the section
   *
   * @protected
   * @memberof MgtPersonCardFiles
   */
  public clearState(): void {
    super.clearState();
    this._files = null;
  }

  /**
   * Render the icon for display in the navigation ribbon.
   *
   * @returns {TemplateResult}
   * @memberof MgtPersonCardFiles
   */
  public renderIcon(): TemplateResult {
    return getSvg(SvgIcon.Files);
  }

  /**
   * Render the compact view
   *
   * @returns {TemplateResult}
   * @memberof MgtPersonCardFiles
   */
  public renderCompactView(): TemplateResult {
    let contentTemplate: TemplateResult;

    if (this.isLoadingState) {
      contentTemplate = this.renderLoading();
    } else if (!this._files || !this._files.length) {
      contentTemplate = this.renderNoData();
    } else {
      contentTemplate = html`
        ${this._files.slice(0, 3).map(file => this.renderFile(file))}
      `;
    }

    return html`
      <div class="root compact" dir=${this.direction}>
        ${contentTemplate}
      </div>
    `;
  }

  /**
   * Render the full view
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCardFiles
   */
  protected renderFullView(): TemplateResult {
    let contentTemplate: TemplateResult;

    if (this.isLoadingState) {
      contentTemplate = this.renderLoading();
    } else if (!this._files || !this._files.length) {
      contentTemplate = this.renderNoData();
    } else {
      contentTemplate = html`
        ${this._files.map(file => this.renderFile(file))}
      `;
    }

    return html`
      <div class="root" dir=${this.direction}>
        <div class="title">${this.strings.filesSectionTitle}</div>
        ${contentTemplate}
      </div>
    `;
  }

  /**
   * Render a file item
   *
   * @protected
   * @param {IFile} file
   * @returns {TemplateResult}
   * @memberof MgtPersonCardFiles
   */
  protected renderFile(file: SharedInsight): TemplateResult {
    const lastModifiedTemplate = file.lastShared
      ? html`
          <div class="file__last-modified">
            ${this.strings.sharedTextSubtitle} ${getRelativeDisplayDate(new Date(file.lastShared.sharedDateTime))}
          </div>
        `
      : null;

    return html`
      <div class="file" @click=${e => this.handleFileClick(file)} tabindex="0">
        <div class="file__icon">
          <img alt="${file.resourceVisualization.title}" src=${getFileTypeIconUri(
      file.resourceVisualization.type,
      48,
      'svg'
    )} />
        </div>
        <div class="file__details">
          <div class="file__name">${file.resourceVisualization.title}</div>
          ${lastModifiedTemplate}
        </div>
      </div>
    `;
  }

  private handleFileClick(file: SharedInsight) {
    if (file.resourceReference && file.resourceReference.webUrl) {
      window.open(file.resourceReference.webUrl, '_blank', 'noreferrer');
    }
  }
}
