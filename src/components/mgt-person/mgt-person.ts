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

    // If we only have a user-id then use that to get data.
    if (this.userId) {
      let person: MgtPersonDetails = {};

      await Promise.all([
        provider.graph.getUser(this.userId).then(user => {
          if (user) {
            person.displayName = user.displayName;
            person.email = user.mail;
            this.requestUpdate();
          }
        }),
        provider.graph.getUserPhoto(this.userId).then(photo => {
          if (photo) {
            person.image = photo;
            this.requestUpdate();
          }
        })
      ]);

      this.personDetails = person;
    } else if (!this.personQuery && this.personDetails) {
      // Check if we have details needed.
      if (!this.personDetails.image) {
        await this.loadImage(this.personDetails);
      }
    } else if (!this.personDetails && this.personQuery) {
      if (this.personQuery == 'me') {
        let person: MgtPersonDetails = {};

        await Promise.all([
          provider.graph.me().then(user => {
            if (user) {
              person.displayName = user.displayName;
              person.email = user.mail || user.userPrincipalName;
            }
          }),
          provider.graph.myPhoto().then(photo => {
            if (photo) {
              person.image = photo;
            }
          })
        ]);

        this.personDetails = person;
      } else {
        provider.graph.findPerson(this.personQuery).then(people => {
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
        });
      }
    } else {
      this.personDetails = null; // Should we throw here or re-load the query as now we do nothing if both specified?
    }
  }

  private async loadImage(person: any) {
    let provider = Providers.globalProvider;

    if (person.userPrincipalName) {
      let userPrincipalName = person.userPrincipalName;
      provider.graph.getUserPhoto(userPrincipalName).then(photo => {
        this.personDetails.image = photo;
        this.requestUpdate();
      });
    } else if (this.personDetails.email) {
      // try to find a user by e-mail
      provider.graph.findUserByEmail(this.personDetails.email).then(users => {
        if (users && users.length) {
          const contactId = users[0].id;
          if ((<any>users[0]).personType && (<any>users[0]).personType.subclass == 'OrganizationUser') {
            provider.graph
              .getUserPhoto((<MicrosoftGraph.Person>users[0]).scoredEmailAddresses[0].address)
              .then(photo => {
                this.personDetails.image = photo;
                this.requestUpdate();
              });
          } else {
            provider.graph.getContactPhoto(contactId).then(photo => {
              this.personDetails.image = photo;
              this.requestUpdate();
            });
          }
        }
      });
    }
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
      if (this.personDetails.image) {
        return html`
          <img
            class="user-avatar ${this.getImageRowSpanClass()} ${this.getImageSizeClass()}"
            title=${this.personDetails.displayName}
            src=${this.personDetails.image as string}
          />
        `;
      } else {
        return html`
          <div
            class="user-avatar initials ${this.getImageRowSpanClass()} ${this.getImageSizeClass()}"
            title=${this.personDetails.displayName}
          >
            <span class="initials-text">
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
