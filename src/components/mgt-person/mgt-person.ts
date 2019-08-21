/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property } from 'lit-element';

import { classMap } from 'lit-html/directives/class-map';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import '../../styles/fabric-icon-font';
import { getEmailFromGraphEntity } from '../../utils/graphHelpers';
import { delay } from '../../utils/utils';
import { MgtTemplatedComponent } from '../templatedComponent';
import { styles } from './mgt-person-css';

@customElement('mgt-person')
export class MgtPerson extends MgtTemplatedComponent {
  static get styles() {
    return styles;
  }
  @property({
    attribute: 'person-query'
  })
  public personQuery: string;

  @property({
    attribute: 'user-id'
  })
  public userId: string;

  @property({
    attribute: 'show-name',
    type: Boolean
  })
  public showName: false;

  @property({
    attribute: 'show-email',
    type: Boolean
  })
  public showEmail: false;

  @property({
    attribute: 'person-details',
    type: Object
  })
  public personDetails: MicrosoftGraph.User | MicrosoftGraph.Person | MicrosoftGraph.Contact;

  @property({
    attribute: 'person-image',
    type: String
  })
  public personImage: string;

  @property({
    attribute: 'show-card',
    type: Boolean
  })
  public showCard: false;

  @property({ attribute: false }) private _isPersonCardVisible: boolean = false;
  @property({ attribute: false }) private _personCardShouldRender: boolean = false;

  private _mouseLeaveTimeout;
  private _mouseEnterTimeout;

  public attributeChangedCallback(name, oldval, newval) {
    super.attributeChangedCallback(name, oldval, newval);

    if ((name === 'person-query' || name === 'user-id') && oldval !== newval) {
      this.personDetails = null;
      this.loadData();
    }
  }

  public firstUpdated() {
    Providers.onProviderUpdated(() => this.loadData());
    this.loadData();
  }

  public render() {
    const person =
      this.renderTemplate('default', { person: this.personDetails }) ||
      html`
        <div class="person-root">
          ${this.renderImage()} ${this.renderDetails()}
        </div>
      `;

    return html`
      <div class="root" @mouseenter=${this._handleMouseEnter} @mouseleave=${this._handleMouseLeave}>
        ${person} ${this.renderPersonCard()}
      </div>
    `;
  }

  private async loadData() {
    const provider = Providers.globalProvider;

    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    if (this.personDetails) {
      // in some cases we might only have name or email, but need to find the image
      // use @ for the image value to search for an image
      if (this.personImage && this.personImage === '@') {
        this.personImage = null;
        this.loadImage();
      }
      return;
    }

    if (this.userId || (this.personQuery && this.personQuery === 'me')) {
      const batch = provider.graph.createBatch();

      if (this.userId) {
        batch.get('user', `/users/${this.userId}`, ['user.readbasic.all']);
        batch.get('photo', `users/${this.userId}/photo/$value`, ['user.readbasic.all']);
      } else {
        batch.get('user', 'me', ['user.read']);
        batch.get('photo', 'me/photo/$value', ['user.read']);
      }

      const response = await batch.execute();

      this.personImage = response.photo;
      this.personDetails = response.user;
    } else if (!this.personDetails && this.personQuery) {
      const people = await provider.graph.findPerson(this.personQuery);
      if (people && people.length > 0) {
        const person = people[0] as MicrosoftGraph.Person;
        this.personDetails = person;

        this.loadImage();
      }
    }
  }

  private async loadImage() {
    const provider = Providers.globalProvider;

    const person = this.personDetails;

    if ((person as MicrosoftGraph.Person).userPrincipalName) {
      const userPrincipalName = (person as MicrosoftGraph.Person).userPrincipalName;
      this.personImage = await provider.graph.getUserPhoto(userPrincipalName);
    } else {
      const email = getEmailFromGraphEntity(person);
      if (email) {
        // try to find a user by e-mail
        const users = await provider.graph.findUserByEmail(email);

        if (users && users.length) {
          if ((users[0] as any).personType && (users[0] as any).personType.subclass === 'OrganizationUser') {
            this.personImage = await provider.graph.getUserPhoto(
              (users[0] as MicrosoftGraph.Person).scoredEmailAddresses[0].address
            );
          } else {
            const contactId = users[0].id;
            this.personImage = await provider.graph.getContactPhoto(contactId);
          }
        }
      }
    }
    this.requestUpdate();
  }

  // todo remove
  private _handleMouseEnter(e: MouseEvent) {
    clearTimeout(this._mouseEnterTimeout);
    clearTimeout(this._mouseLeaveTimeout);
    this._mouseEnterTimeout = setTimeout(this._showPersonCard.bind(this), 300, e);
  }

  private _handleMouseLeave(e: MouseEvent) {
    clearTimeout(this._mouseEnterTimeout);
    clearTimeout(this._mouseLeaveTimeout);
    this._mouseLeaveTimeout = setTimeout(this._hidePersonCard.bind(this), 500, e);
  }

  private async _showPersonCard(e: MouseEvent) {
    if (!this._personCardShouldRender) {
      this._personCardShouldRender = true;
    }

    // give the person-card a chance to render so transitions work
    await delay(200);

    this._isPersonCardVisible = true;
  }

  private _hidePersonCard(e: MouseEvent) {
    this._isPersonCardVisible = false;
  }

  private renderPersonCard() {
    // ensure person card is only rendered when needed
    if (!this.showCard || !this._personCardShouldRender) {
      return;
    }

    const flyoutClasses = { flyout: true, visible: this._isPersonCardVisible };
    return html`
      <div class=${classMap(flyoutClasses)}>
        <mgt-person-card .person=></mgt-person-card>
      </div>
    `;
  }

  private renderDetails() {
    if (this.showEmail || this.showName) {
      return html`
        <span class="Details ${this.getImageSizeClass()}">
          ${this.renderNameAndEmail()}
        </span>
      `;
    }

    return null;
  }

  private renderImage() {
    if (this.personDetails) {
      const title = !this.showCard ? this.personDetails.displayName : '';

      if (this.personImage && this.personImage !== '@') {
        return html`
          <img
            class="user-avatar ${this.getImageRowSpanClass()} ${this.getImageSizeClass()}"
            title=${title}
            src=${this.personImage as string}
          />
        `;
      } else {
        return html`
          <div class="user-avatar initials ${this.getImageRowSpanClass()} ${this.getImageSizeClass()}" title=${title}>
            <span class="initials-text">
              ${this.getInitials()}
            </span>
          </div>
        `;
      }
    }

    return this.renderEmptyImage();
  }

  private renderEmptyImage() {
    return html`
      <i class="ms-Icon ms-Icon--Contact avatar-icon ${this.getImageRowSpanClass()} ${this.getImageSizeClass()}"></i>
    `;
  }

  private renderNameAndEmail() {
    if (!this.personDetails || (!this.showEmail && !this.showName)) {
      return;
    }

    const nameView = this.showName
      ? html`
          <div class="user-name">${this.personDetails.displayName}</div>
        `
      : null;
    const emailView = this.showEmail
      ? html`
          <div class="user-email">${getEmailFromGraphEntity(this.personDetails)}</div>
        `
      : null;

    return html`
      ${nameView} ${emailView}
    `;
  }

  private getInitials() {
    if (!this.personDetails) {
      return '';
    }

    let initials = '';
    if (this.personDetails.givenName) {
      initials += this.personDetails.givenName[0].toUpperCase();
    }
    if (this.personDetails.surname) {
      initials += this.personDetails.surname[0].toUpperCase();
    }

    if (!initials && this.personDetails.displayName) {
      const name = this.personDetails.displayName.split(' ');
      for (let i = 0; i < 2 && i < name.length; i++) {
        if (name[i][0].match(/[a-z]/i)) {
          // check if letter
          initials += name[i][0].toUpperCase();
        }
      }
    }

    return initials;
  }

  private getImageRowSpanClass() {
    if (this.showEmail && this.showName) {
      return 'row-span-2';
    }

    return '';
  }

  private getImageSizeClass() {
    if (!this.showEmail || !this.showName) {
      return 'small';
    }

    return '';
  }
}
