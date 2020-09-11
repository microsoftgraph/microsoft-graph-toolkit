/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, TemplateResult } from 'lit-element';
import { getMe } from '../../../../graph/graph.user';
import { Providers, ProviderState } from '@microsoft/mgt-element';
import { getRelativeDisplayDate } from '../../../../utils/Utils';
import { BasePersonCardSection } from '../BasePersonCardSection';
import { getMessages, getMessagesWithUser, IMessage } from './graph.messages';
import { styles } from './mgt-person-card-messages-css';
import { getEmailFromGraphEntity } from '../../../../graph/graph.people';

/**
 * The email messages subsection of the person card
 *
 * @export
 * @class MgtPersonCardMessages
 * @extends {MgtTemplatedComponent}
 */
@customElement('mgt-person-card-messages')
export class MgtPersonCardMessages extends BasePersonCardSection {
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
   * @memberof MgtPersonCardMessages
   */
  public get displayName(): string {
    return 'Emails';
  }

  private _messages: IMessage[];

  /**
   * Reset any state in the section
   *
   * @protected
   * @memberof MgtPersonCardMessages
   */
  public clearState(): void {
    this._messages = [];
  }

  /**
   * Render the icon for display in the navigation ribbon.
   *
   * @returns {TemplateResult}
   * @memberof MgtPersonCardMessages
   */
  public renderIcon(): TemplateResult {
    return html`
      <svg xmlns="http://www.w3.org/2000/svg">
        <path
          fill-rule="evenodd"
          clip-rule="evenodd"
          d="M4 3.5C4 3.22386 4.22386 3 4.5 3H16C17.1046 3 18 3.89543 18 5V12.5C18 12.7761 17.7761 13 17.5 13C17.2239 13 17 12.7761 17 12.5V5C17 4.44772 16.5523 4 16 4H4.5C4.22386 4 4 3.77614 4 3.5ZM4.04886 6H13.8473L9.04316 9.19968C8.86969 9.31522 8.64273 9.31103 8.47364 9.18916L4.04886 6ZM3 14V6.47671L7.88894 10.0004C8.39621 10.366 9.07706 10.3786 9.59749 10.032L15 6.43376V14H3ZM3 5C2.44772 5 2 5.44772 2 6V14C2 14.5523 2.44772 15 3 15H15C15.5523 15 16 14.5523 16 14V6C16 5.44772 15.5523 5 15 5H3Z"
        />
      </svg>
    `;
  }

  /**
   * Render the compact view
   *
   * @returns {TemplateResult}
   * @memberof MgtPersonCardMessages
   */
  public renderCompactView(): TemplateResult {
    let contentTemplate: TemplateResult;

    if (this.isLoadingState) {
      contentTemplate = this.renderLoading();
    } else if (!this._messages || !this._messages.length) {
      contentTemplate = this.renderNoData();
    } else {
      const messageTemplates = this._messages
        ? this._messages.slice(0, 3).map(message => this.renderMessage(message))
        : [];
      contentTemplate = html`
        ${messageTemplates}
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
   * @memberof MgtPersonCardMessages
   */
  protected renderFullView(): TemplateResult {
    let contentTemplate: TemplateResult;

    if (this.isLoadingState) {
      contentTemplate = this.renderLoading();
    } else if (!this._messages || !this._messages.length) {
      contentTemplate = this.renderNoData();
    } else {
      contentTemplate = html`
        ${this._messages.slice(0, 5).map(message => this.renderMessage(message))}
      `;
    }

    return html`
      <div class="root">
        <div class="title">Emails</div>
        ${contentTemplate}
      </div>
    `;
  }

  /**
   * Render a message item
   *
   * @protected
   * @param {IMessage} message
   * @returns {TemplateResult}
   * @memberof MgtPersonCardMessages
   */
  protected renderMessage(message: IMessage): TemplateResult {
    return html`
      <div class="message" @click=${() => this.handleMessageClick(message)}>
        <div class="message__detail">
          <div class="message__subject">${message.subject}</div>
          <div class="message__from">${message.from.emailAddress.name}</div>
          <div class="message__message">${message.bodyPreview}</div>
        </div>
        <div class="message__date">${getRelativeDisplayDate(new Date(message.receivedDateTime))}</div>
      </div>
    `;
  }

  private handleMessageClick(message: IMessage): void {
    window.open(message.webLink, '_blank');
  }

  /**
   * load state into the component
   *
   * @protected
   * @returns {Promise<void>}
   * @memberof MgtPersonCardMessages
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
    const me = await getMe(graph);

    const userId = this.personDetails.id;

    if (me.id === userId) {
      this._messages = await getMessages(graph);
    } else {
      const emailAddress = getEmailFromGraphEntity(this.personDetails);
      if (emailAddress) {
        this._messages = await getMessagesWithUser(graph, emailAddress);
      }
    }

    this.requestUpdate();
  }
}
