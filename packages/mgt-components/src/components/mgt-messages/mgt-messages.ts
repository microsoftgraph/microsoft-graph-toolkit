/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Message } from '@microsoft/microsoft-graph-types';
import { html, TemplateResult } from 'lit';

import { BasePersonCardSection } from '../BasePersonCardSection';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';
import { getRelativeDisplayDate } from '../../utils/Utils';
import { styles } from './mgt-messages-css';
import { strings } from './strings';
import { customElement } from '@microsoft/mgt-element';

/**
 * The email messages subsection of the person card
 *
 * @export
 * @class MgtMessages
 * @extends {MgtTemplatedComponent}
 */
@customElement('messages')
export class MgtMessages extends BasePersonCardSection {
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
   * @memberof MgtMessages
   */
  public get displayName(): string {
    return this.strings.emailsSectionTitle;
  }

  /**
   * The title for display when rendered as a full card.
   *
   * @readonly
   * @type {string}
   * @memberof MgtOrganization
   */
  public get cardTitle(): string {
    return this.strings.emailsSectionTitle;
  }

  /**
   * Reset any state in the section
   *
   * @protected
   * @memberof MgtMessages
   */
  public clearState(): void {
    super.clearState();
    this._messages = [];
  }

  /**
   * Render the icon for display in the navigation ribbon.
   *
   * @returns {TemplateResult}
   * @memberof MgtMessages
   */
  public renderIcon(): TemplateResult {
    return getSvg(SvgIcon.Messages);
  }

  /**
   * Render the compact view
   *
   * @returns {TemplateResult}
   * @memberof MgtMessages
   */
  public renderCompactView(): TemplateResult {
    let contentTemplate: TemplateResult;

    if (this.isLoadingState) {
      contentTemplate = this.renderLoading();
    } else if (!this._messages?.length) {
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
   * @memberof MgtMessages
   */
  protected renderFullView(): TemplateResult {
    let contentTemplate: TemplateResult;

    if (this.isLoadingState) {
      contentTemplate = this.renderLoading();
    } else if (!this._messages?.length) {
      contentTemplate = this.renderNoData();
    } else {
      contentTemplate = html`
         ${this._messages.slice(0, 5).map(message => this.renderMessage(message))}
       `;
    }

    return html`
       <div class="root">
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
   * @memberof MgtMessages
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
    window.open(message.webLink, '_blank', 'noreferrer');
  }
}
