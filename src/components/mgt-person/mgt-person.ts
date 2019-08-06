/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html, customElement, property } from 'lit-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import { Providers } from '../../Providers';
import { styles } from './mgt-person-css';
import '../../styles/fabric-icon-font';
import { ProviderState } from '../../providers/IProvider';
import { MgtTemplatedComponent } from '../templatedComponent';
import { classMap } from 'lit-html/directives/class-map';
import { delay } from '../../utils/utils';

@customElement('mgt-person')
export class MgtPerson extends MgtTemplatedComponent {
  @property({
    attribute: 'person-query'
  })
  personQuery: string;

  @property({
    attribute: 'user-id'
  })
  userId: string;

  @property({
    attribute: 'show-name',
    type: Boolean
  })
  showName: false;

  @property({
    attribute: 'show-email',
    type: Boolean
  })
  showEmail: false;

  @property({
    attribute: 'person-details',
    type: Object
  })
  personDetails: MgtPersonDetails;

  @property({
    attribute: 'show-card',
    type: Boolean
  })
  showCard: false;

  @property({ attribute: false }) private _isPersonCardVisible: boolean = false;
  @property({ attribute: false }) private _personCardHasRendered: boolean = false;

  attributeChangedCallback(name, oldval, newval) {
    super.attributeChangedCallback(name, oldval, newval);

    if ((name == 'person-query' || name == 'user-id') && oldval !== newval) {
      this.personDetails = null;
      this.loadData();
    }
  }

  static get styles() {
    return styles;
  }

  firstUpdated() {
    Providers.onProviderUpdated(() => this.loadData());
    this.loadData();
  }

  private async loadData() {
    let provider = Providers.globalProvider;

    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    if (this.personDetails) {
      // in some cases we might only have name or email, but need to find the image
      // use @ for the image value to search for an image
      if (this.personDetails.image && this.personDetails.image === '@') {
        this.personDetails.image = null;
        this.loadImage(this.personDetails);
      }
      return;
    }

    if (this.userId || (this.personQuery && this.personQuery == 'me')) {
      let batch = provider.graph.createBatch();

      if (this.userId) {
        batch.get('user', `/users/${this.userId}`, ['user.readbasic.all']);
        batch.get('photo', `users/${this.userId}/photo/$value`, ['user.readbasic.all']);
      } else {
        batch.get('user', 'me', ['user.read']);
        batch.get('photo', 'me/photo/$value', ['user.read']);
      }

      let response = await batch.execute();

      this.personDetails = {
        displayName: response.user.displayName,
        email: response.user.mail || response.user.userPrincipalName,
        image: response.photo
      };
    } else if (!this.personDetails && this.personQuery) {
      let people = await provider.graph.findPerson(this.personQuery);
      if (people && people.length > 0) {
        let person = people[0] as MicrosoftGraph.Person;
        this.personDetails = person;

        if (person.scoredEmailAddresses && person.scoredEmailAddresses.length) {
          this.personDetails.email = person.scoredEmailAddresses[0].address;
        } else if ((<any>person).emailAddresses && (<any>person).emailAddresses.length) {
          // beta endpoint uses emailAddresses instead of scoredEmailAddresses
          this.personDetails.email = (<any>person).emailAddresses[0].address;
        }

        this.loadImage(person);
      }
    }
  }

  private async loadImage(person: any) {
    let provider = Providers.globalProvider;

    if (person.userPrincipalName) {
      let userPrincipalName = person.userPrincipalName;
      this.personDetails.image = await provider.graph.getUserPhoto(userPrincipalName);
    } else if (this.personDetails.email) {
      // try to find a user by e-mail
      let users = await provider.graph.findUserByEmail(this.personDetails.email);

      if (users && users.length) {
        if ((<any>users[0]).personType && (<any>users[0]).personType.subclass == 'OrganizationUser') {
          this.personDetails.image = await provider.graph.getUserPhoto(
            (<MicrosoftGraph.Person>users[0]).scoredEmailAddresses[0].address
          );
        } else {
          const contactId = users[0].id;
          this.personDetails.image = await provider.graph.getContactPhoto(contactId);
        }
      }
    }
    this.requestUpdate();
  }

  private _mouseLeaveTimeout;
  private _mouseEnterTimeout;

  // todo remove
  private _handleMouseEnter(e: MouseEvent) {
    clearTimeout(this._mouseEnterTimeout);
    clearTimeout(this._mouseLeaveTimeout);
    this._mouseEnterTimeout = setTimeout(this._showPersonCard.bind(this), 1000, e);
  }

  private _handleMouseLeave(e: MouseEvent) {
    clearTimeout(this._mouseEnterTimeout);
    clearTimeout(this._mouseLeaveTimeout);
    this._mouseLeaveTimeout = setTimeout(this._hidePersonCard.bind(this), 1000, e);
  }

  private async _showPersonCard(e: MouseEvent) {
    if (!this._personCardHasRendered) {
      // give the person-card a chance to render so transitions work
      this._personCardHasRendered = true;
      await delay(200);
    }

    this._isPersonCardVisible = true;
  }

  private _hidePersonCard(e: MouseEvent) {
    this._isPersonCardVisible = false;
  }

  render() {
    let person =
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

  private renderPersonCard() {
    // ensure person card is only rendered when needed
    if (!this.showCard || !this._personCardHasRendered) {
      return;
    }

    let flyoutClasses = { flyout: true, visible: this._isPersonCardVisible };
    return html`
      <div class=${classMap(flyoutClasses)}>
        <mgt-person-card></mgt-person-card>
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
      let title = !this.showCard ? this.personDetails.displayName : '';

      if (this.personDetails.image && this.personDetails.image !== '@') {
        return html`
          <img
            class="user-avatar ${this.getImageRowSpanClass()} ${this.getImageSizeClass()}"
            title=${title}
            src=${this.personDetails.image as string}
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
          <div class="user-email">${this.personDetails.email}</div>
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

export declare interface MgtPersonDetails {
  displayName?: string;
  email?: string;
  image?: string;
  givenName?: string;
  surname?: string;
}
