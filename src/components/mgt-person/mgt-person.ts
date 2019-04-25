import { LitElement, html, customElement, property } from 'lit-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import { Providers } from '../../Providers';
import { styles } from './mgt-person-css';
import '../../styles/fabric-icon-font';
import { ProviderState } from '../../providers/IProvider';
import { MgtTemplatedComponent } from '../templatedComponent';

@customElement('mgt-person')
export class MgtPerson extends MgtTemplatedComponent {
  @property({
    attribute: 'image-size'
  })
  imageSize: number = 24;

  @property({
    attribute: 'person-query'
  })
  personQuery: string;

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

    if (name == 'person-query' && oldval !== newval) {
      this.personDetails = null;
      this.loadImage();
    }
  }

  static get styles() {
    return styles;
  }

  constructor() {
    super();
    Providers.onProviderUpdated(() => this.loadImage());
    this.loadImage();
  }

  private async loadImage() {
    if (!this.personDetails && this.personQuery) {
      let provider = Providers.globalProvider;

      if (provider && provider.state === ProviderState.SignedIn) {
        if (this.personQuery == 'me') {
          let person: MgtPersonDetails = {};

          await Promise.all([
            provider.graph.me().then(user => {
              if (user) {
                person.displayName = user.displayName;
                person.email = user.mail;
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
                // beta endpoind uses emailAddresses instead of scoredEmailAddresses
                this.personDetails.email = (<any>person).emailAddresses[0].address;
              }

              if (person.userPrincipalName) {
                let userPrincipalName = person.userPrincipalName;
                provider.graph.getUserPhoto(userPrincipalName).then(photo => {
                  this.personDetails.image = photo;
                  this.requestUpdate();
                });
              }
            }
          });
        }
      } else {
        this.personDetails = null;
      }
    }
  }

  render() {
    let templates = this.getTemplates();
    this.removeSlottedElements();

    if (templates['default']) {
      return this.renderTemplate(
        templates['default'],
        {
          person: this.personDetails
        },
        'default'
      );
    }

    return html`
      <div class="root">
        ${this.renderImage()} ${this.renderNameAndEmail()}
      </div>
    `;
  }

  renderImage() {
    if (this.personDetails) {
      if (this.personDetails.image) {
        return html`
          <img
            class="user-avatar ${this.getImageRowSpanClass()} ${this.getImageSizeClass()}"
            src=${this.personDetails.image as string}
          />
        `;
      } else {
        return html`
          <div class="user-avatar initials ${this.getImageRowSpanClass()} ${this.getImageSizeClass()}">
            <style>
              .initials-text {
                font-size: ${this.imageSize * 0.45}px;
              }
            </style>
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
        initials += name[i][0].toUpperCase();
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
