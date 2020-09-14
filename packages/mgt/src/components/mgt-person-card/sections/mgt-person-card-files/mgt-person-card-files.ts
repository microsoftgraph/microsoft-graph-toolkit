/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, TemplateResult } from 'lit-element';
import { BetaGraph } from '../../../../BetaGraph';
import { Providers, ProviderState } from '@microsoft/mgt-element';
import { getRelativeDisplayDate } from '../../../../utils/Utils';
import { BasePersonCardSection } from '../BasePersonCardSection';
import { IFile, getFilesSharedWithMe, getFilesSharedWithUser, getThumbnails } from './graph.files';
import { styles } from './mgt-person-card-files-css';
import { getEmailFromGraphEntity } from '../../../../graph/graph.people';
import { getMe } from '../../../../graph/graph.user';
import { getSvg, SvgIcon } from '../../../../utils/SvgHelper';

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

  /**
   * The name for display in the overview section.
   *
   * @readonly
   * @type {string}
   * @memberof MgtPersonCardFiles
   */
  public get displayName(): string {
    return 'Files';
  }

  private _files: IFile[];

  /**
   * Reset any state in the section
   *
   * @protected
   * @memberof MgtPersonCardFiles
   */
  public clearState(): void {
    this._files = [];
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
      <div class="root compact">
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
      <div class="root">
        <div class="title">Files</div>
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
  protected renderFile(file: IFile): TemplateResult {
    const iconSource = file.thumbnails ? file.thumbnails[0].small.url : null;

    const lastModifiedTemplate = file.lastModifiedDateTime
      ? html`
          <div class="file__last-modified">Modified ${getRelativeDisplayDate(new Date(file.lastModifiedDateTime))}</div>
        `
      : null;

    return html`
      <div class="file">
        <div class="file__icon">
          <img src="${iconSource}" />
        </div>
        <div class="file__details">
          <div class="file__name">${file.name}</div>
          ${lastModifiedTemplate}
        </div>
      </div>
    `;
  }

  /**
   * load state into the component
   *
   * @protected
   * @returns {Promise<void>}
   * @memberof MgtPersonCardProfile
   */
  protected async loadState(): Promise<void> {
    const provider = Providers.globalProvider;

    // check if user is signed in
    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    if (!this.personDetails) {
      return;
    }

    const graph = provider.graph.forComponent(this);
    const betaGraph = BetaGraph.fromGraph(graph);

    const me = await getMe(graph);
    const userId = this.personDetails.id;

    if (me.id === userId) {
      this._files = await getFilesSharedWithMe(betaGraph);
    } else {
      const emailAddress = getEmailFromGraphEntity(this.personDetails);
      if (emailAddress) {
        this._files = await getFilesSharedWithUser(betaGraph, emailAddress);
      }
    }

    if (this._files) {
      for (const file of this._files) {
        file.thumbnails = await getThumbnails(graph, file);
      }
    }

    this.requestUpdate();
  }
}
