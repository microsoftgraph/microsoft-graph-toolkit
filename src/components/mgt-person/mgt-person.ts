/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import { customElement, html, property, query, TemplateResult } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { findPeople, getEmailFromGraphEntity } from '../../graph/graph.people';
import { getPersonImage } from '../../graph/graph.photos';
import { getUserPresence } from '../../graph/graph.presence';
import { getUserWithPhoto } from '../../graph/graph.user';
import { AvatarSize, IDynamicPerson } from '../../graph/types';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import '../../styles/fabric-icon-font';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';
import { MgtPersonCard } from '../mgt-person-card/mgt-person-card';
import '../sub-components/mgt-flyout/mgt-flyout';
import { MgtFlyout } from '../sub-components/mgt-flyout/mgt-flyout';
import { MgtTemplatedComponent } from '../templatedComponent';
import { PersonCardInteraction } from './../PersonCardInteraction';
import { styles } from './mgt-person-css';

export { PersonCardInteraction } from './../PersonCardInteraction';

/**
 * The person component is used to display a person or contact by using their photo, name, and/or email address.
 *
 * @export
 * @class MgtPerson
 * @extends {MgtTemplatedComponent}
 *
 * @cssprop --avatar-size - {Length} Avatar size
 * @cssprop --avatar-border - {String} Avatar border
 * @cssprop --initials-color - {Color} Initials color
 * @cssprop --initials-background-color - {Color} Initials background color
 * @cssprop --font-size - {Length} Font size
 * @cssprop --font-weight - {Length} Font weight
 * @cssprop --default-font-family - {String} Font family
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
   * determines if person component renders presence
   * @type {boolean}
   */
  @property({
    attribute: 'show-presence',
    type: Boolean
  })
  public showPresence: boolean;

  /**
   * determines person component avatar size and apply presence badge accordingly
   * @type {AvatarSize}
   */
  @property({
    attribute: 'avatar-size',
    type: String
  })
  public avatarSize: AvatarSize = 'unset';

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

    if (this.fetchImage) {
      this.personImage = null;
    }

    if (this.showPresence) {
      this.personPresence = null;
    }

    this.requestStateUpdate();
    this.requestUpdate('personDetails');
  }

  /**
   * Set the image of the person
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
   * Sets whether the person image should be fetched
   * from the Microsoft Graph based on the personDetails
   * provided by the user
   *
   * @type {boolean}
   * @memberof MgtPerson
   */
  @property({
    attribute: 'fetch-image',
    type: Boolean
  })
  public fetchImage: boolean;

  /**
   * Gets or sets presence of person
   *
   * @type {MicrosoftGraphBeta.Presence}
   * @memberof MgtPerson
   */
  @property({
    attribute: 'person-presence',
    type: Object
  })
  public personPresence: MicrosoftGraphBeta.Presence;

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
    const presence = this.personPresence;

    if (!person && !image) {
      return this.renderNoData();
    }

    // Default template
    let personTemplate = this.renderTemplate('default', { person, personImage: image });

    if (!personTemplate) {
      const detailsTemplate: TemplateResult = this.renderDetails(person);
      const imageWithPresenceTemplate: TemplateResult = this.renderImageWithPresence(image, person, presence);

      personTemplate = html`
        <div class="person-root">
          ${imageWithPresenceTemplate} ${detailsTemplate}
        </div>
      `;
    }

    if (this.personCardInteraction !== PersonCardInteraction.none) {
      personTemplate = this.renderFlyout(personTemplate);
    }

    return html`
      <div
        class="root"
        @click=${this.handleMouseClick}
        @mouseenter=${this.handleMouseEnter}
        @mouseleave=${this.handleMouseLeave}
      >
        ${personTemplate}
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

    const isLarge = this.avatarSize === 'large' || (this.avatarSize === 'unset' && this.showEmail && this.showName);

    const avatarClasses = {
      'avatar-icon': true,
      'ms-Icon': true,
      'ms-Icon--Contact': true,
      small: !isLarge
    };

    return html`
      <i class=${classMap(avatarClasses)}></i>
    `;
  }

  /**
   * Render the image part of the person template.
   * If the image is unavailable, the person's initials will be used instead.
   *
   * @protected
   * @param {string} [imageSrc]
   * @param {IDynamicPerson} [personDetails]
   * @returns
   * @memberof MgtPerson
   */
  protected renderImage(imageSrc: string, personDetails: IDynamicPerson) {
    const title =
      personDetails && this.personCardInteraction === PersonCardInteraction.none ? personDetails.displayName : '';

    if (imageSrc) {
      return html`
        <img alt=${title} src=${imageSrc} />
      `;
    } else if (personDetails) {
      const initials = this.getInitials(personDetails);

      return html`
        <span class="initials-text" aria-label="${initials}">
          ${initials && initials.length
            ? html`
                ${initials}
              `
            : html`
                <i class="ms-Icon ms-Icon--Contact contact-icon"></i>
              `}
        </span>
      `;
    }
  }

  /**
   * Render presence for the person.
   *
   * @protected
   * @param
   * @memberof MgtPersonCard
   */
  protected renderPresence(presence: MicrosoftGraphBeta.Presence): TemplateResult {
    if (!this.showPresence || !presence) {
      return html``;
    }

    let statusClass = null;
    // attach appropriate css class to show different icons
    switch (presence.availability) {
      case 'DoNotDisturb':
        switch (presence.activity) {
          case 'OutOfOffice':
            statusClass = 'presence-oof-dnd';
            break;
          default:
            statusClass = 'presence-dnd';
            break;
        }
        break;
      case 'BeRightBack':
        statusClass = 'presence-away';
        break;
      case 'Available':
        switch (presence.activity) {
          case 'Available':
            statusClass = 'presence-available';
            break;
          case 'OutOfOffice':
            statusClass = 'presence-oof-available';
            break;
        }
        break;
      case 'Busy':
        switch (presence.activity) {
          case 'OutOfOffice':
            statusClass = 'presence-oof-busy';
            break;
          default:
            // 'Busy', 'InACall', 'InAMeeting'
            statusClass = 'presence-busy';
            break;
        }
        break;
      case 'Away':
        switch (presence.activity) {
          case 'Away':
            statusClass = 'presence-away';
            break;
          case 'OutOfOffice':
            statusClass = 'presence-oof-offline';
            break;
        }
        break;
      case 'Offline':
        switch (presence.activity) {
          case 'Offline':
            statusClass = 'presence-offline';
            break;
          case 'OutOfOffice':
            statusClass = 'presence-oof-offline';
            break;
        }
        break;
      default:
        statusClass = 'presence-offline';
        break;
    }

    const presenceClasses = {
      'ms-Icon': true,
      'presence-basic': true
    };

    presenceClasses[statusClass] = true;
    // workaround because SkypeArrow icon from fluent doesn't work ¯\_(ツ)_/¯
    let iconHtml = null;
    if (statusClass === 'presence-oof-offline') {
      iconHtml = html`
        <div class="ms-Icon presence-basic presence-oof-offline-wrapper">
          <i class="presence-oof-offline">
            ${getSvg(SvgIcon.SkypeArrow, '#666666')}
          </i>
        </div>
      `;
    } else {
      iconHtml = html`
        <i class=${classMap(presenceClasses)} aria-hidden="true"></i>
      `;
    }

    return html`
      <div class="user-presence">
        ${iconHtml}
      </div>
    `;
  }

  /**
   * Render image with presence for the person.
   *
   * @protected
   * @param
   * @memberof MgtPersonCard
   */
  protected renderImageWithPresence(
    image: string,
    personDetails: IDynamicPerson,
    presence: MicrosoftGraphBeta.Presence
  ): TemplateResult {
    const title =
      this.personDetails && this.personCardInteraction === PersonCardInteraction.none
        ? this.personDetails.displayName
        : '';

    const isLarge = this.avatarSize === 'large' || (this.avatarSize === 'unset' && this.showEmail && this.showName);
    const imageClasses = {
      initials: !image,
      small: !isLarge,
      'user-avatar': true
    };

    if (!image && this.personDetails) {
      // add avatar background color
      imageClasses[this._personAvatarBg] = true;
    }

    const imageTemplate: TemplateResult = this.renderImage(image, personDetails);
    const presenceTemplate: TemplateResult = this.renderPresence(presence);

    return html`
      <div class=${classMap(imageClasses)} title=${title} aria-label=${title}>
        ${imageTemplate} ${presenceTemplate}
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
  protected renderDetails(person: IDynamicPerson): TemplateResult {
    if (!person || (!this.showEmail && !this.showName)) {
      return html``;
    }

    const email = getEmailFromGraphEntity(person);
    const emailTemplate: TemplateResult = this.showEmail && !!email ? this.renderEmail(email) : html``;
    const nameTemplate: TemplateResult = this.showName ? this.renderName(person.displayName) : html``;

    const isLarge = this.avatarSize === 'large' || (this.avatarSize === 'unset' && this.showEmail && this.showName);

    const detailsClasses = classMap({
      details: true,
      small: !isLarge
    });

    return html`
      <div class="${detailsClasses}">
        ${nameTemplate} ${emailTemplate}
      </div>
    `;
  }

  /**
   * Render the name part of the person details.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPerson
   */
  protected renderName(displayName: string): TemplateResult {
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
  protected renderEmail(email: string): TemplateResult {
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
    const personPresence = this.personPresence;
    const showPresence = this.showPresence;
    return (
      this.renderTemplate('person-card', { person, personImage: image }) ||
      html`
        <mgt-person-card
          .personDetails=${person}
          .personImage=${image}
          .personPresence=${personPresence}
          .showPresence=${showPresence}
        ></mgt-person-card>
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
      if (!this.personDetails.personImage && ((this.fetchImage && !this.personImage) || this.personImage === '@')) {
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
      this.personDetails.personImage = this.personImage;
    } else if (this.personQuery) {
      // Use the personQuery to find our person.
      const people = await findPeople(graph, this.personQuery, 1);

      if (people && people.length) {
        this.personDetails = people[0];
        const image = await getPersonImage(graph, people[0]);

        if (image) {
          this.personDetails.personImage = image;
          this.personImage = image;
        }
      }
    }

    // populate presence
    const defaultPresence = {
      activity: 'Offline',
      availability: 'Offline',
      id: null
    };
    if (!this.personPresence && this.showPresence) {
      try {
        if (this.personDetails && this.personDetails.id) {
          this.personPresence = await getUserPresence(graph, this.personDetails.id);
        } else {
          this.personPresence = defaultPresence;
        }
      } catch (_) {
        // set up a default Presence in case beta api changes or getting error code
        this.personPresence = defaultPresence;
      }
    }

    // populate avatar size
    if (this.showEmail && this.showName) {
      this.avatarSize = 'large';
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
        if (name[i][0] && this.isLetter(name[i][0])) {
          initials += name[i][0].toUpperCase();
        }
      }
    }

    return initials;
  }

  private isLetter(char: string) {
    try {
      return char.match(new RegExp('\\p{L}', 'u'));
    } catch (e) {
      return char.toLowerCase() !== char.toUpperCase();
    }
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
