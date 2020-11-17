/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, TemplateResult } from 'lit-element';
import { getRelativeDisplayDate } from '../../../../utils/Utils';
import { BasePersonCardSection } from '../BasePersonCardSection';
import { styles } from './mgt-person-card-messages-css';
import { SvgIcon, getSvg } from '../../../../utils/SvgHelper';
import { Message } from '@microsoft/microsoft-graph-types';

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

  private _messages: Message[];

  public constructor(messages: Message[]) {
    super();
    this._messages = messages;
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
    return getSvg(SvgIcon.Messages);
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
  protected renderMessage(message: Message): TemplateResult {
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

  private handleMessageClick(message: Message): void {
    window.open(message.webLink, '_blank');
  }
}
