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
import { IFile, getFilesSharedWithMe, getFilesSharedWithUser } from './graph.files';
import { styles } from './mgt-person-card-files-css';
import { getEmailFromGraphEntity } from '../../../../graph/graph.people';
import { getMe } from '../../../../graph/graph.user';

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
    return html`
      <svg width="20" height="20" viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg">
        <path
          d="M17 8C17.1354 8 17.263 8.02604 17.3828 8.07812C17.5078 8.13021 17.6146 8.20312 17.7031 8.29688C17.7969 8.38542 17.8698 8.48958 17.9219 8.60938C17.974 8.72917 18 8.85677 18 8.99219C18 9.14844 17.9635 9.29948 17.8906 9.44531L14.6172 16H2V5C2 4.85938 2.02604 4.72917 2.07812 4.60938C2.13021 4.48958 2.20052 4.38542 2.28906 4.29688C2.38281 4.20312 2.48958 4.13021 2.60938 4.07812C2.72917 4.02604 2.85938 4 3 4H5.75C5.98438 4 6.1849 4.02604 6.35156 4.07812C6.52344 4.125 6.67448 4.1849 6.80469 4.25781C6.9401 4.33073 7.0599 4.41146 7.16406 4.5C7.26823 4.58854 7.3724 4.66927 7.47656 4.74219C7.58594 4.8151 7.70052 4.8776 7.82031 4.92969C7.94531 4.97656 8.08854 5 8.25 5H14C14.1406 5 14.2708 5.02604 14.3906 5.07812C14.5104 5.13021 14.6146 5.20312 14.7031 5.29688C14.7969 5.38542 14.8698 5.48958 14.9219 5.60938C14.974 5.72917 15 5.85938 15 6V8H17ZM3 13.3828L5.41406 8.55469C5.5026 8.38281 5.625 8.2474 5.78125 8.14844C5.94271 8.04948 6.11979 8 6.3125 8H14V6H8.25C8.01562 6 7.8125 5.97656 7.64062 5.92969C7.47396 5.8776 7.32292 5.8151 7.1875 5.74219C7.05729 5.66927 6.9401 5.58854 6.83594 5.5C6.73177 5.41146 6.625 5.33073 6.51562 5.25781C6.41146 5.1849 6.29688 5.125 6.17188 5.07812C6.05208 5.02604 5.91146 5 5.75 5H3V13.3828ZM17 9H6.3125L3.3125 15H14L17 9Z"
        />
      </svg>
    `;
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
    const iconTemplate = html`
      <img src="${1}" />
    `;

    const lastModifiedTemplate = file.lastModifiedDateTime
      ? html`
          <div class="file__last-modified">Modified ${getRelativeDisplayDate(new Date(file.lastModifiedDateTime))}</div>
        `
      : null;

    return html`
      <div class="file">
        <div class="file__icon">${iconTemplate}</div>
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
        const thumbnailRequest = graph.api(`/users/${userId}/drive/items/${file.id}/thumbnails`).get();
        thumbnailRequest
          .then(thumbnails => {
            file.thumbnails = thumbnails;
          })
          .catch(e => {
            console.log(e);
          });
      }
    }

    this.requestUpdate();
  }
}
