/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property, PropertyValues } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import '../../styles/fabric-icon-font';
import { getEmailFromGraphEntity } from '../../utils/GraphHelpers';
import { MgtPersonCard } from '../mgt-person-card/mgt-person-card';
import { MgtTemplatedComponent } from '../templatedComponent';
import { PersonCardInteraction } from './../PersonCardInteraction';
import { styles } from './mgt-person-css';

/**
 * The person component is used to display a person or contact by using their photo, name, and/or email address.
 *
 * @export
 * @class MgtPerson
 * @extends {MgtTemplatedComponent}
 */
@customElement('mgt-person')
export class MgtPerson extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * user the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * allows developer to define name of person for component
   * @type {string}
   */
  @property({
    attribute: 'person-query'
  })
  public personQuery: string;

  /**
   * user-id property allows developer to use id value to determine person
   * @type {string}
   */
  @property({
    attribute: 'user-id'
  })
  public userId: string;

  /**
   * determines if person component renders user-name
   * @type {boolean}
   */
  @property({
    attribute: 'show-name',
    type: Boolean
  })
  public showName: false;

  /**
   * determines if person component renders email
   * @type {boolean}
   */
  @property({
    attribute: 'show-email',
    type: Boolean
  })
  public showEmail: false;

  /**
   * object containing Graph details on person
   * @type {MgtPersonDetails}
   */
  @property({
    attribute: 'person-details',
    type: Object
  })
  public personDetails: MicrosoftGraph.User | MicrosoftGraph.Person | MicrosoftGraph.Contact;

  /**
   * Set the image of the person
   * Set to '@' to look up image from the graph
   *
   * @type {string}
   * @memberof MgtPersonCard
   */
  @property({
    attribute: 'person-image',
    type: String
  })
  public personImage: string;

  /**
   * Sets how the person-card is invoked
   * Set to PersonCardInteraction.none to not show the card
   *
   * @type {PersonCardInteraction}
   * @memberof MgtPerson
   */
  @property({
    attribute: 'person-card',
    converter: (value, type) => {
      value = value.toLowerCase();
      if (typeof PersonCardInteraction[value] === 'undefined') {
        return PersonCardInteraction.none;
      } else {
        return PersonCardInteraction[value];
      }
    }
  })
  public personCardInteraction: PersonCardInteraction = PersonCardInteraction.none;

  @property({ attribute: false }) private _isPersonCardVisible: boolean = false;
  @property({ attribute: false }) private _personCardShouldRender: boolean = false;

  private _mouseLeaveTimeout;
  private _mouseEnterTimeout;
  private _openLeft: boolean = false;
  private _openUp: boolean = false;

  /**
   * Synchronizes property values when attributes change.
   *
   * @param {*} name
   * @param {*} oldValue
   * @param {*} newValue
   * @memberof MgtPerson
   */
  public attributeChangedCallback(name, oldval, newval) {
    super.attributeChangedCallback(name, oldval, newval);

    if ((name === 'person-query' || name === 'user-id') && oldval !== newval) {
      this.personDetails = null;
      this.loadData();
    }
  }

  /**
   * Invoked when the element is first updated. Implement to perform one time
   * work on the element after update.
   *
   * Setting properties inside this method will trigger the element to update
   * again after this update cycle completes.
   *
   * * @param _changedProperties Map of changed properties with old values
   */

  public firstUpdated() {
    Providers.onProviderUpdated(() => this.loadData());
    this.loadData();
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  public render() {
    const person =
      this.renderTemplate('default', { person: this.personDetails, personImage: this.getImage() }) ||
      html`
        <div class="person-root">
          ${this.renderImage()} ${this.renderDetails()}
        </div>
      `;

    return html`
      <div
        class="root"
        @mouseenter=${this._handleMouseEnter}
        @mouseleave=${this._handleMouseLeave}
        @click=${this._handleMouseClick}
      >
        ${person} ${this.renderPersonCard()}
      </div>
    `;
  }

  /**
   * Invoked whenever the element is updated. Implement to perform
   * post-updating tasks via DOM APIs, for example, focusing an element.
   *
   * Setting properties inside this method will trigger the element to update
   * again after this update cycle completes.
   *
   * * @param changedProperties Map of changed properties with old values
   */
  protected updated(changedProps: PropertyValues) {
    super.updated(changedProps);

    const initials = this.renderRoot.querySelector('.initials-text') as HTMLElement;
    if (initials && initials.parentNode && (initials.parentNode as HTMLElement).getBoundingClientRect) {
      const parent = initials.parentNode as HTMLElement;
      const height = parent.getBoundingClientRect().height;
      initials.style.fontSize = `${height * 0.5}px`;
    }
  }

  private async loadData() {
    const provider = Providers.globalProvider;

    if (!provider || provider.state === ProviderState.Loading) {
      return;
    }

    if (provider.state === ProviderState.SignedOut) {
      this.personDetails = null;
      return;
    }

    // personDetails.personImage is a toolkit injected property to pass image between components
    // an optimization to avoid fetching the image when unnecessary
    if (this.personDetails) {
      // in some cases we might only have name or email, but need to find the image
      // use @ for the image value to search for an image
      if (this.personImage && this.personImage === '@' && !(this.personDetails as any).personImage) {
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

      this.personDetails = response.user;
      (this.personDetails as any).personImage = response.photo;
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
    let image: string;

    if ((person as MicrosoftGraph.Person).userPrincipalName) {
      const userPrincipalName = (person as MicrosoftGraph.Person).userPrincipalName;
      image = await provider.graph.getUserPhoto(userPrincipalName);
    } else {
      const email = getEmailFromGraphEntity(person);
      if (email) {
        // try to find a user by e-mail
        const users = await provider.graph.findUserByEmail(email);

        if (users && users.length) {
          if ((users[0] as any).personType && (users[0] as any).personType.subclass === 'OrganizationUser') {
            image = await provider.graph.getUserPhoto(
              (users[0] as MicrosoftGraph.Person).scoredEmailAddresses[0].address
            );
          } else {
            const contactId = users[0].id;
            image = await provider.graph.getContactPhoto(contactId);
          }
        }
      }
    }
    if (image) {
      (this.personDetails as any).personImage = image;
    }

    this.requestUpdate();
  }

  private _handleMouseClick() {
    if (this.personCardInteraction === PersonCardInteraction.click && !this._isPersonCardVisible) {
      this._showPersonCard();
    } else {
      this._hidePersonCard();
    }
  }

  private _handleMouseEnter(e: MouseEvent) {
    if (this.personCardInteraction !== PersonCardInteraction.hover) {
      return;
    }

    clearTimeout(this._mouseEnterTimeout);
    clearTimeout(this._mouseLeaveTimeout);
    this._mouseEnterTimeout = setTimeout(this._showPersonCard.bind(this), 500);
  }

  private _handleMouseLeave(e: MouseEvent) {
    clearTimeout(this._mouseEnterTimeout);
    clearTimeout(this._mouseLeaveTimeout);
    this._mouseLeaveTimeout = setTimeout(this._hidePersonCard.bind(this), 500);
  }

  private _showPersonCard() {
    if (!this._personCardShouldRender) {
      this._personCardShouldRender = true;
    }

    this._isPersonCardVisible = true;
  }

  private _hidePersonCard() {
    this._isPersonCardVisible = false;
    const personCard = this.querySelector('mgt-person-card') as MgtPersonCard;
    if (personCard) {
      personCard.isExpanded = false;
    }
  }

  private getImage(): string {
    if (this.personImage && this.personImage !== '@') {
      return this.personImage;
    } else if (this.personDetails && (this.personDetails as any).personImage) {
      return (this.personDetails as any).personImage;
    }
    return null;
  }

  private renderPersonCard() {
    // ensure person card is only rendered when needed
    if (this.personCardInteraction === PersonCardInteraction.none || !this._personCardShouldRender) {
      return;
    }
    // logic for rendering left if there is no space
    const personRect = this.renderRoot.querySelector('.root').getBoundingClientRect();
    const leftEdge = personRect.left;
    const rightEdge = (window.innerWidth || document.documentElement.clientWidth) - personRect.right;
    this._openLeft = rightEdge < leftEdge;

    // logic for rendering up
    const bottomEdge = (window.innerHeight || document.documentElement.clientHeight) - personRect.bottom;
    this._openUp = bottomEdge < 175;

    // find position to renderup to
    let customStyle = null;
    if (this._openUp) {
      const personSize = this.getBoundingClientRect().bottom - this.getBoundingClientRect().top;
      customStyle = `bottom: ${personSize / 2 + 8}px`;
    }

    const flyoutClasses = {
      flyout: true,
      openLeft: this._openLeft,
      openUp: this._openUp,
      visible: this._isPersonCardVisible
    };
    if (this._isPersonCardVisible) {
      const image = this.getImage();

      return html`
        <div style="${customStyle}" class=${classMap(flyoutClasses)}>
          ${this.renderTemplate('person-card', { person: this.personDetails, personImage: image }) ||
            html`
              <mgt-person-card .personDetails=${this.personDetails} .personImage=${image}> </mgt-person-card>
            `}
        </div>
      `;
    }
  }

  private renderDetails() {
    if (this.showEmail || this.showName) {
      const isLarge = this.showEmail && this.showName;

      const detailsClasses = {
        Details: true,
        small: !isLarge
      };

      return html`
        <span class="${classMap(detailsClasses)}">
          ${this.renderNameAndEmail()}
        </span>
      `;
    }

    return null;
  }

  private renderImage() {
    if (this.personDetails) {
      const title = this.personCardInteraction === PersonCardInteraction.none ? this.personDetails.displayName : '';
      const image = this.getImage();
      const isLarge = this.showEmail && this.showName;
      const imageClasses = {
        initials: !image,
        'row-span-2': isLarge,
        small: !isLarge,
        'user-avatar': true
      };

      let imageHtml;

      if (image) {
        imageHtml = html`
          <img alt=${title} src=${image} />
        `;
      } else {
        const initials = this.getInitials();

        imageHtml = html`
          <span class="initials-text" aria-label="${initials}">
            ${initials}
          </span>
        `;
      }

      return html`
        <div class=${classMap(imageClasses)} title=${title} aria-label=${title}>
          ${imageHtml}
        </div>
      `;
    }

    return this.renderEmptyImage();
  }

  private renderEmptyImage() {
    const isLarge = this.showEmail && this.showName;

    const imageClasses = {
      'avatar-icon': true,
      'ms-Icon': true,
      'ms-Icon--Contact': true,
      'row-span-2': isLarge,
      small: !isLarge
    };

    return html`
      <i class=${classMap(imageClasses)}></i>
    `;
  }

  private renderNameAndEmail() {
    if (!this.personDetails || (!this.showEmail && !this.showName)) {
      return;
    }

    const nameView = this.showName
      ? html`
          <div class="user-name" aria-label="${this.personDetails.displayName}">${this.personDetails.displayName}</div>
        `
      : null;

    let emailView;
    if (this.showEmail) {
      const email = getEmailFromGraphEntity(this.personDetails);
      emailView = html`
        <div class="user-email" aria-label="${email}">${email}</div>
      `;
    }

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
}
