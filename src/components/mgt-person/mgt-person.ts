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

@customElement('mgt-person')
export class MgtPerson extends MgtTemplatedComponent {
  /**
   * person-query property, allows developer to define name of person for component
   * @type {string}
   */
  @property({
    attribute: 'person-query'
  })
  personQuery: string;

  /**
   * user-id property, allows developer to use id value to determine person
   * @type {string}
   */
  @property({
    attribute: 'user-id'
  })
  userId: string;

  /**
   * show-name property, determines if person component renders user-name
   * @type {boolean}
   */
  @property({
    attribute: 'show-name',
    type: Boolean
  })
  showName: false;

  /**
   * show-email property, determines if person component renders email
   * @type {boolean}
   */
  @property({
    attribute: 'show-email',
    type: Boolean
  })
  showEmail: false;

  /**
   * person-details property, object containing Graph details on person
   * @type {MgtPersonDetails}
   */
  @property({
    attribute: 'person-details',
    type: Object
  })
  personDetails: MgtPersonDetails;

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

  render() {
    return (
      this.renderTemplate('default', { person: this.personDetails }) ||
      html`
        <div class="root">
          ${this.renderImage()} ${this.renderDetails()}
        </div>
      `
    );
  }

  renderDetails() {
    if (this.showEmail || this.showName) {
      return html`
        <span class="Details ${this.getImageSizeClass()}">
          ${this.renderNameAndEmail()}
        </span>
      `;
    }

    return null;
  }

  renderImage() {
    if (this.personDetails) {
      if (this.personDetails.image && this.personDetails.image !== '@') {
        return html`
          <img
            class="user-avatar ${this.getImageRowSpanClass()} ${this.getImageSizeClass()}"
            title=${this.personDetails.displayName}
            aria-label=${this.personDetails.displayName}
            alt=${this.personDetails.displayName}
            src=${this.personDetails.image as string}
          />
        `;
      } else {
        return html`
          <div
            class="user-avatar initials ${this.getImageRowSpanClass()} ${this.getImageSizeClass()}"
            title=${this.personDetails.displayName}
            aria-label=${this.personDetails.displayName}
          >
            <span class="initials-text" aria-label="${this.getInitials()}">
              ${this.getInitials()}
            </span>
          </div>
        `;
      }
    }

    return this.renderEmptyImage();
  }

  renderEmptyImage() {
    return html`
      <i class="ms-Icon ms-Icon--Contact avatar-icon ${this.getImageRowSpanClass()} ${this.getImageSizeClass()}"></i>
    `;
  }

  renderNameAndEmail() {
    if (!this.personDetails || (!this.showEmail && !this.showName)) {
      return;
    }

    const nameView = this.showName
      ? html`
          <div class="user-name" aria-label="${this.personDetails.displayName}">${this.personDetails.displayName}</div>
        `
      : null;
    const emailView = this.showEmail
      ? html`
          <div class="user-email" aria-label="${this.personDetails.email}">${this.personDetails.email}</div>
        `
      : null;

    return html`
      ${nameView} ${emailView}
    `;
  }

  getInitials() {
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

  getImageRowSpanClass() {
    if (this.showEmail && this.showName) {
      return 'row-span-2';
    }

    return '';
  }

  getImageSizeClass() {
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
