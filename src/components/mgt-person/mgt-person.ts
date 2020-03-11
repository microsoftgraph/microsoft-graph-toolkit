/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property, PropertyValues, TemplateResult } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { findPerson, getEmailFromGraphEntity } from '../../graph/graph.people';
import { getContactPhoto, getUserPhoto } from '../../graph/graph.photos';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import '../../styles/fabric-icon-font';
import { MgtPersonCard } from '../mgt-person-card/mgt-person-card';
import '../sub-components/mgt-flyout/mgt-flyout';
import { MgtTemplatedComponent } from '../templatedComponent';
import { PersonCardInteraction } from './../PersonCardInteraction';
import { styles } from './mgt-person-css';
import { findUserByEmail } from './mgt-person.graph';

/**
 * The person component is used to display a person or contact by using their photo, name, and/or email address.
 *
 * @export
 * @class MgtPerson
 * @extends {MgtTemplatedComponent}
 *
 * @cssprop --avatar-size-s - {Length} Avatar size
 * @cssprop --avatar-size - {Length} Avatar size when both name and email are shown
 * @cssprop --avatar-font-size--s - {Length} Avatar font size
 * @cssprop --avatar-font-size - {Length} Avatar font-size when both name and email are shown
 * @cssprop --avatar-border - {String} Avatar border
 * @cssprop --initials-color - {Color} Initials color
 * @cssprop --initials-background-color - {Color} Initials background color
 * @cssprop --font-size - {Length} Font size
 * @cssprop --font-weight - {Length} Font weight
 * @cssprop --color - {Color} Color
 * @cssprop --email-font-size - {Length} Email font size
 * @cssprop --email-color - {Color} Email color
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
  public showName: boolean;

  /**
   * determines if person component renders email
   * @type {boolean}
   */
  @property({
    attribute: 'show-email',
    type: Boolean
  })
  public showEmail: boolean;

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
    reflect: true,
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
  public personCardInteraction: PersonCardInteraction;

  @property({ attribute: false }) private isPersonCardVisible: boolean;
  @property({ attribute: false }) private personCardShouldRender: boolean;

  private _mouseLeaveTimeout;
  private _mouseEnterTimeout;

  constructor() {
    super();
    this.personCardInteraction = PersonCardInteraction.none;
  }

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

    if (oldval === newval) {
      return;
    }

    switch (name) {
      case 'person-query':
      case 'user-id':
        this.personDetails = null;
        this.requestStateUpdate();
        break;
    }
  }

  /**
   * Invoked each time the custom element is appended into a document-connected element
   *
   * @memberof MgtPerson
   */
  public connectedCallback() {
    super.connectedCallback();
    window.addEventListener('click', e => this.handleWindowClick(e));
  }

  /**
   * Invoked each time the custom element is disconnected from the document's DOM
   *
   * @memberof MgtPerson
   */
  public disconnectedCallback() {
    window.removeEventListener('click', e => this.handleWindowClick(e));
    super.disconnectedCallback();
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  public render() {
    // Loading
    if (this.isLoadingState) {
      return this.renderLoading();
    }

    // No data
    if (!this.personDetails) {
      return this.renderNoData();
    }

    // Prep data
    const image = this.getImage();

    // Default template
    const renderedTemplate = this.renderTemplate('default', { person: this.personDetails, personImage: image });
    if (renderedTemplate) {
      return renderedTemplate;
    }

    // Person template
    const personTemplate = this.renderPerson(this.personDetails, image);

    return html`
      <div
        class="root"
        @click=${this.handleMouseClick}
        @mouseenter=${this.handleMouseEnter}
        @mouseleave=${this.handleMouseLeave}
      >
        ${this.personCardInteraction === PersonCardInteraction.none
          ? personTemplate
          : this.renderFlyout(personTemplate)}
      </div>
    `;
  }

  /**
   * foo
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPerson
   */
  protected renderLoading(): TemplateResult {
    return this.renderTemplate('loading', null) || html``;
  }

  /**
   * foo
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPerson
   */
  protected renderNoData(): TemplateResult {
    if (!this.renderTemplate('no-data', null)) {
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
  }

  /**
   * foo
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPerson
   */
  protected renderPerson(
    person: MicrosoftGraph.User | MicrosoftGraph.Person | MicrosoftGraph.Contact,
    image: string
  ): TemplateResult {
    const imageTemplate: TemplateResult = this.renderImage(image);
    const detailsTemplate: TemplateResult = this.renderDetails(person);

    return (
      this.renderTemplate('person', { person, personImage: image }) ||
      html`
        <div class="person-root">
          ${imageTemplate} ${detailsTemplate}
        </div>
      `
    );
  }

  /**
   * foo
   *
   * @protected
   * @param {string} image
   * @returns
   * @memberof MgtPerson
   */
  protected renderImage(image: string) {
    const title = this.personCardInteraction === PersonCardInteraction.none ? this.personDetails.displayName : '';
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

  /**
   * foo
   *
   * @protected
   * @param {(MicrosoftGraph.User | MicrosoftGraph.Person | MicrosoftGraph.Contact)} person
   * @returns
   * @memberof MgtPerson
   */
  protected renderDetails(person: MicrosoftGraph.User | MicrosoftGraph.Person | MicrosoftGraph.Contact) {
    if (!this.showEmail && !this.showName) {
      return html``;
    }

    let emailTemplate: TemplateResult;
    if (this.showEmail) {
      const email = getEmailFromGraphEntity(person);
      emailTemplate = this.renderEmail(email);
    } else {
      emailTemplate = html``;
    }

    const nameTemplate: TemplateResult = this.showName ? this.renderName(person.displayName) : html``;

    const isLarge = this.showEmail && this.showName;
    const detailsClasses = classMap({
      Details: true,
      small: !isLarge
    });

    return html`
      <span class="${detailsClasses}">
        ${nameTemplate} ${emailTemplate}
      </span>
    `;
  }

  /**
   * foo
   *
   * @protected
   * @memberof MgtPerson
   */
  protected renderName(displayName: string): TemplateResult {
    return html`
      <div class="user-name" aria-label="${displayName}">${displayName}</div>
    `;
  }

  /**
   * foo
   *
   * @protected
   * @memberof MgtPerson
   */
  protected renderEmail(email: string): TemplateResult {
    return html`
      <div class="user-email" aria-label="${email}">${email}</div>
    `;
  }

  /**
   * foo
   *
   * @protected
   * @param {TemplateResult} anchor
   * @returns
   * @memberof MgtPerson
   */
  protected renderFlyout(anchor: TemplateResult) {
    const person = this.personDetails;
    const image = this.getImage();
    const flyout = this.personCardShouldRender
      ? html`
          <div slot="flyout" class="flyout">
            ${this.renderPersonCard(person, image)}
          </div>
        `
      : null;

    return html`
      <mgt-flyout .isOpen=${this.isPersonCardVisible}>
        ${anchor} ${flyout}
      </mgt-flyout>
    `;
  }

  /**
   * foo
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPerson
   */
  protected renderPersonCard(
    person: MicrosoftGraph.User | MicrosoftGraph.Person | MicrosoftGraph.Contact,
    image: string
  ): TemplateResult {
    return (
      this.renderTemplate('person-card', { person, personImage: image }) ||
      html`
        <mgt-person-card .personDetails=${person} .personImage=${image}></mgt-person-card>
      `
    );
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

  /**
   * load state into the component.
   *
   * @protected
   * @returns
   * @memberof MgtPerson
   */
  protected async loadState() {
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
      const graph = provider.graph.forComponent(this);
      const batch = graph.createBatch();

      if (this.userId) {
        batch.get('user', `/users/${this.userId}`, ['user.readbasic.all']);
        batch.get('photo', `users/${this.userId}/photo/$value`, ['user.readbasic.all']);
      } else {
        batch.get('user', 'me', ['user.read']);
        batch.get('photo', 'me/photo/$value', ['user.read']);
      }

      const response = await batch.execute();

      this.personDetails = response.user;
      this.personImage = response.photo;
      (this.personDetails as any).personImage = response.photo;
    } else if (!this.personDetails && this.personQuery) {
      const graph = provider.graph.forComponent(this);
      const people = await findPerson(graph, this.personQuery);
      if (people && people.length > 0) {
        const person = people[0] as MicrosoftGraph.Person;
        this.personDetails = person;

        this.loadImage();
      }
    }
  }

  private async loadImage() {
    const provider = Providers.globalProvider;
    const graph = provider.graph.forComponent(this);

    const person = this.personDetails;
    let image: string;

    if ((person as MicrosoftGraph.Person).userPrincipalName) {
      const userPrincipalName = (person as MicrosoftGraph.Person).userPrincipalName;
      image = await getUserPhoto(graph, userPrincipalName);
    } else {
      const email = getEmailFromGraphEntity(person);
      if (email) {
        // try to find a user by e-mail
        const users = await findUserByEmail(graph, email);

        if (users && users.length) {
          if ((users[0] as any).personType && (users[0] as any).personType.subclass === 'OrganizationUser') {
            image = await getUserPhoto(graph, (users[0] as MicrosoftGraph.Person).scoredEmailAddresses[0].address);
          } else {
            const contactId = users[0].id;
            image = await getContactPhoto(graph, contactId);
          }
        }
      }
    }
    if (image) {
      this.personImage = image;
      (this.personDetails as any).personImage = image;
    }

    this.requestUpdate();
  }

  private getImage(): string {
    if (this.personImage && this.personImage !== '@') {
      return this.personImage;
    } else if (this.personDetails && (this.personDetails as any).personImage) {
      return (this.personDetails as any).personImage;
    }
    return null;
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

  private handleWindowClick(e: MouseEvent) {
    if (this.isPersonCardVisible && e.target !== this) {
      this.hidePersonCard();
    }
  }

  private handleMouseClick(e: MouseEvent) {
    if (this.personCardInteraction !== PersonCardInteraction.none && !this.isPersonCardVisible) {
      this.showPersonCard();
    }
  }

  private handleMouseEnter(e: MouseEvent) {
    clearTimeout(this._mouseEnterTimeout);
    clearTimeout(this._mouseLeaveTimeout);
    if (this.personCardInteraction !== PersonCardInteraction.hover) {
      return;
    }
    this._mouseEnterTimeout = setTimeout(this.showPersonCard.bind(this), 500);
  }

  private handleMouseLeave(e: MouseEvent) {
    clearTimeout(this._mouseEnterTimeout);
    clearTimeout(this._mouseLeaveTimeout);
    this._mouseLeaveTimeout = setTimeout(this.hidePersonCard.bind(this), 500);
  }

  private hidePersonCard() {
    this.isPersonCardVisible = false;
    const personCard = (this.querySelector('mgt-person-card') ||
      this.renderRoot.querySelector('mgt-person-card')) as MgtPersonCard;
    if (personCard) {
      personCard.isExpanded = false;
    }
  }

  private showPersonCard() {
    if (!this.personCardShouldRender) {
      this.personCardShouldRender = true;
    }

    this.isPersonCardVisible = true;
  }
}
