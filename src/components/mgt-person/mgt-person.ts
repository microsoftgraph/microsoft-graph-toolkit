/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property, query, TemplateResult } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { findPerson, getEmailFromGraphEntity } from '../../graph/graph.people';
import { getPersonImage } from '../../graph/graph.photos';
import { getUserWithPhoto } from '../../graph/graph.user';
import { IDynamicPerson } from '../../graph/types';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import '../../styles/fabric-icon-font';
import { MgtPersonCard } from '../mgt-person-card/mgt-person-card';
import '../sub-components/mgt-flyout/mgt-flyout';
import { MgtFlyout } from '../sub-components/mgt-flyout/mgt-flyout';
import { MgtTemplatedComponent } from '../templatedComponent';
import { PersonCardInteraction } from './../PersonCardInteraction';
import { styles } from './mgt-person-css';

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
   * @type {IDynamicPerson}
   */
  @property({
    attribute: 'person-details',
    type: Object
  })
  public get personDetails(): IDynamicPerson {
    return this._personDetails;
  }

  public set personDetails(value: IDynamicPerson) {
    if (this._personDetails === value) {
      return;
    }

    this._personDetails = value;
    if (!value) {
      this._personAvatarBg = 'gray20';
    } else {
      this._personAvatarBg = this.getColorFromName(value.displayName);
    }
    this.requestUpdate('personDetails');
  }

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

  /**
   * Gets the flyout element
   *
   * @protected
   * @type {MgtFlyout}
   * @memberof MgtPerson
   */
  protected get flyout(): MgtFlyout {
    return this.renderRoot.querySelector('.flyout');
  }

  @property({ attribute: false }) private _personCardShouldRender: boolean;
  private _personDetails: IDynamicPerson;
  private _personAvatarBg: string;

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
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  public render() {
    // Loading
    if (this.isLoadingState && !this.personDetails) {
      return this.renderLoading();
    }

    // Prep data
    const person = this.personDetails;
    const image = this.getImage();

    // Default template
    let personTemplate = this.renderTemplate('default', { person, personImage: image });
    if (!personTemplate) {
      const imageTemplate: TemplateResult = this.renderImage(image);
      const detailsTemplate: TemplateResult = this.renderDetails(person);
      personTemplate = html`
        <div class="person-root">
          ${imageTemplate} ${detailsTemplate}
        </div>
      `;
    }

    // Flyout template
    const flyoutTemplate: TemplateResult =
      this.personCardInteraction !== PersonCardInteraction.none ? this.renderFlyout(personTemplate) : null;

    return html`
      <div
        class="root"
        @click=${this.handleMouseClick}
        @mouseenter=${this.handleMouseEnter}
        @mouseleave=${this.handleMouseLeave}
      >
        ${flyoutTemplate || personTemplate}
      </div>
    `;
  }

  /**
   * Render the loading state
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPerson
   */
  protected renderLoading(): TemplateResult {
    return this.renderTemplate('loading', null) || html``;
  }

  /**
   * Render the state when no data is available
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPerson
   */
  protected renderNoData(): TemplateResult {
    const noDataTemplate = this.renderTemplate('no-data', null);
    if (noDataTemplate) {
      return noDataTemplate;
    }

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

  /**
   * Render the image part of the person template.
   * If the image is unavailable, the person's initials will be used instead.
   *
   * @protected
   * @param {string} [imageSrc]
   * @returns
   * @memberof MgtPerson
   */
  protected renderImage(imageSrc?: string) {
    if (!imageSrc) {
      imageSrc = this.getImage();
    }

    const title =
      this.personDetails && this.personCardInteraction === PersonCardInteraction.none
        ? this.personDetails.displayName
        : '';
    const isLarge = this.showEmail && this.showName;
    const imageClasses = {
      initials: !imageSrc,
      'row-span-2': isLarge,
      small: !isLarge,
      'user-avatar': true
    };

    let imageHtml: TemplateResult;
    if (imageSrc) {
      // render the image
      imageHtml = html`
        <img alt=${title} src=${imageSrc} />
      `;
    } else if (this.personDetails) {
      // render the initials
      const initials = this.getInitials(this.personDetails);
      // add avatar background color
      imageClasses[this._personAvatarBg] = true;
      imageHtml = html`
        <span class="initials-text" aria-label="${initials}">
          ${initials}
        </span>
      `;
    } else {
      return this.renderNoData();
    }

    return html`
      <div class=${classMap(imageClasses)} title=${title} aria-label=${title}>
        ${imageHtml}
      </div>
    `;
  }

  /**
   * Render the details part of the person template.
   *
   * @protected
   * @param {IDynamicPerson} [person]
   * @param {string} [image]
   * @returns {TemplateResult}
   * @memberof MgtPerson
   */
  protected renderDetails(person?: IDynamicPerson): TemplateResult {
    if (!this.showEmail && !this.showName) {
      return html``;
    }

    person = person || this.personDetails;
    if (!person) {
      return html``;
    }

    const email = getEmailFromGraphEntity(person);
    const emailTemplate: TemplateResult = this.showEmail ? this.renderEmail(email) : html``;
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
   * Render the name part of the person details.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPerson
   */
  protected renderName(displayName?: string): TemplateResult {
    if (!displayName && this.personDetails) {
      displayName = this.personDetails.displayName;
    }
    return html`
      <div class="user-name" aria-label="${displayName}">${displayName}</div>
    `;
  }

  /**
   * Render the email part of the person details.
   *
   * @protected
   * @memberof MgtPerson
   */
  protected renderEmail(email?: string): TemplateResult {
    if (!email) {
      email = getEmailFromGraphEntity(this.personDetails);
    }
    return html`
      <div class="user-email" aria-label="${email}">${email}</div>
    `;
  }

  /**
   * Render the details flyout.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPerson
   */
  protected renderFlyout(anchor: TemplateResult): TemplateResult {
    const flyoutContent = this._personCardShouldRender
      ? html`
          <div slot="flyout">
            ${this.renderFlyoutContent()}
          </div>
        `
      : html``;

    return html`
      <mgt-flyout light-dismiss class="flyout">
        ${anchor} ${flyoutContent}
      </mgt-flyout>
    `;
  }

  /**
   * Render the flyout menu content.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPerson
   */
  protected renderFlyoutContent(): TemplateResult {
    const person = this.personDetails;
    const image = this.getImage();
    return (
      this.renderTemplate('person-card', { person, personImage: image }) ||
      html`
        <mgt-person-card .personDetails=${person} .personImage=${image}></mgt-person-card>
      `
    );
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

    const graph = provider.graph.forComponent(this);
    if (this.personDetails) {
      // in some cases we might only have name or email, but need to find the image
      // use @ for the image value to search for an image
      if (this.personImage === '@' && !this.personDetails.personImage) {
        const image = await getPersonImage(graph, this.personDetails);
        if (image) {
          this.personDetails.personImage = image;
          this.personImage = image;
        }
      }
    } else if (this.userId || this.personQuery === 'me') {
      // Use userId or 'me' query to get the person and image
      const person = await getUserWithPhoto(graph, this.userId);

      this.personDetails = person;
      this.personImage = this.getImage();
    } else if (this.personQuery) {
      // Use the personQuery to find our person.
      const people = await findPerson(graph, this.personQuery);

      if (people && people.length) {
        this.personDetails = people[0];
        const image = await getPersonImage(graph, people[0]);

        if (image) {
          this.personDetails.personImage = image;
          this.personImage = image;
        }
      }
    }
  }

  private getImage(): string {
    if (this.personImage && this.personImage !== '@') {
      return this.personImage;
    }

    const person = this.personDetails;
    return person && person.personImage ? person.personImage : null;
  }

  private getInitials(person?: IDynamicPerson) {
    if (!person) {
      person = this.personDetails;
    }

    if ((person as MicrosoftGraph.Contact).initials) {
      return (person as MicrosoftGraph.Contact).initials;
    }

    let initials = '';
    if (person.givenName) {
      initials += person.givenName[0].toUpperCase();
    }
    if (person.surname) {
      initials += person.surname[0].toUpperCase();
    }

    if (!initials && person.displayName) {
      const name = person.displayName.split(/\s+/);
      for (let i = 0; i < 2 && i < name.length; i++) {
        if (name[i][0] && name[i][0].match(/[a-z]/i)) {
          // check if letter
          initials += name[i][0].toUpperCase();
        }
      }
    }

    return initials;
  }

  private handleMouseClick(e: MouseEvent) {
    if (this.personCardInteraction === PersonCardInteraction.click) {
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
    const flyout = this.flyout;
    if (flyout) {
      flyout.isOpen = false;
    }

    const personCard = (this.querySelector('mgt-person-card') ||
      this.renderRoot.querySelector('mgt-person-card')) as MgtPersonCard;
    if (personCard) {
      personCard.isExpanded = false;
    }
  }

  private showPersonCard() {
    if (!this._personCardShouldRender) {
      this._personCardShouldRender = true;
    }

    const flyout = this.flyout;
    if (flyout) {
      flyout.isOpen = true;
    }
  }

  private getColorFromName(name) {
    const charCodes = name
      .split('')
      .map(char => char.charCodeAt(0))
      .join('');
    const nameInt = parseInt(charCodes, 10);
    const colors = [
      'pinkRed10',
      'red20',
      'red10',
      'orange20',
      'orangeYellow20',
      'green10',
      'green20',
      'cyan20',
      'cyan30',
      'cyanBlue10',
      'cyanBlue20',
      'blue10',
      'blueMagenta30',
      'blueMagenta20',
      'magenta20',
      'magenta10',
      'magentaPink10',
      'orange30',
      'gray30',
      'gray20'
    ];
    return colors[nameInt % colors.length];
  }
}
