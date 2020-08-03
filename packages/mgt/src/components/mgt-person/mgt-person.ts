/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import { customElement, html, internalProperty, property, TemplateResult } from 'lit-element';
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

export { PersonCardInteraction } from '../PersonCardInteraction';

/**
 * Enumeration to define what parts of the person component render
 *
 * @export
 * @enum {number}
 */
export enum PersonViewType {
  /**
   * Render only the avatar
   */
  avatar = 2,

  /**
   * Render the avatar and one line of text
   */
  oneline = 3,

  /**
   * Render the avatar and two lines of text
   */
  twolines = 4
}

/**
 * The person component is used to display a person or contact by using their photo, name, and/or email address.
 *
 * @export
 * @class MgtPerson
 * @extends {MgtTemplatedComponent}
 *
 * @cssprop --avatar-size - {Length} Avatar size
 * @cssprop --avatar-border - {String} Avatar border
 * @cssprop --avatar-border-radius - {String} Avatar border radius
 * @cssprop --initials-color - {Color} Initials color
 * @cssprop --initials-background-color - {Color} Initials background color
 * @cssprop --font-family - {String} Font family
 * @cssprop --font-size - {Length} Font size
 * @cssprop --font-weight - {Length} Font weight
 * @cssprop --color - {Color} Color
 * @cssprop --presence-background-color - {Color} Presence badge background color
 * @cssprop --text-transform - {String} text transform
 * @cssprop --line2-font-size - {Length} Line 2 font size
 * @cssprop --line2-font-weight - {Length} Line 2 font weight
 * @cssprop --line2-color - {Color} Line 2 color
 * @cssprop --line2-text-transform - {String} Line 2 text transform
 * @cssprop --details-spacing - {Length} spacing between avatar and person details
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
   * @deprecated [This property will be removed in next major update. Use `view` instead]
   */
  @property({
    attribute: 'show-name',
    type: Boolean
  })
  public showName: boolean;

  /**
   * determines if person component renders email
   * @type {boolean}
   * @deprecated [This property will be removed in next major update. Use `view` instead]
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
  public avatarSize: AvatarSize;

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
    if (value && value.displayName) {
      this._personAvatarBg = this.getColorFromName(value.displayName);
    } else {
      this._personAvatarBg = 'gray20';
    }

    this._fetchedImage = null;
    this._fetchedPresence = null;

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
  public get personImage(): string {
    return this._personImage || this._fetchedImage;
  }
  public set personImage(value: string) {
    if (value === this._personImage) {
      return;
    }

    this._isInvalidImageSrc = !value;
    const oldValue = this._personImage;
    this._personImage = value;
    this.requestUpdate('personImage', oldValue);
  }

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
  public get personPresence(): MicrosoftGraphBeta.Presence {
    return this._personPresence || this._fetchedPresence;
  }
  public set personPresence(value: MicrosoftGraphBeta.Presence) {
    if (value === this._personPresence) {
      return;
    }

    const oldValue = this._personPresence;
    this._personPresence = value;
    this.requestUpdate('personPresence', oldValue);
  }

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

  /**
   * Sets the property of the personDetails to use for the first line of text.
   * Default is displayName.
   *
   * @type {string}
   * @memberof MgtPerson
   */
  @property({ attribute: 'line1-property' }) public line1Property: string;

  /**
   * Sets the property of the personDetails to use for the second line of text.
   * Default is mail.
   *
   * @type {string}
   * @memberof MgtPerson
   */
  @property({ attribute: 'line2-property' }) public line2Property: string;

  /**
   * Sets what data to be rendered (avatar only, oneLine, twoLines).
   * Default is 'avatar'.
   *
   * @type {PersonViewType}
   * @memberof MgtPerson
   */
  @property({
    converter: value => {
      if (!value || value.length === 0) {
        return PersonViewType.avatar;
      }

      value = value.toLowerCase();

      if (typeof PersonViewType[value] === 'undefined') {
        return PersonViewType.avatar;
      } else {
        return PersonViewType[value];
      }
    }
  })
  public view: PersonViewType;

  @internalProperty() private _fetchedImage: string;
  @internalProperty() private _fetchedPresence: MicrosoftGraphBeta.Presence;
  @internalProperty() private _isInvalidImageSrc: boolean;
  @internalProperty() private _personCardShouldRender: boolean;

  private _personDetails: IDynamicPerson;
  private _personAvatarBg: string;
  private _personImage: string;
  private _personPresence: MicrosoftGraphBeta.Presence;

  private _mouseLeaveTimeout;
  private _mouseEnterTimeout;

  constructor() {
    super();

    // defaults
    this.personCardInteraction = PersonCardInteraction.none;
    this.line1Property = 'displayName';
    this.line2Property = 'email';
    this.view = PersonViewType.avatar;
    this.avatarSize = 'auto';
    this._isInvalidImageSrc = false;
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
    const presence = this.personPresence || this._fetchedPresence;

    if (!person && !image) {
      return this.renderNoData();
    }

    // Default template
    let personTemplate = this.renderTemplate('default', { person, personImage: image });

    if (!personTemplate) {
      const detailsTemplate: TemplateResult = this.renderDetails(person);
      const imageWithPresenceTemplate: TemplateResult = this.renderAvatar(person, image, presence);

      personTemplate = html`
        <div class="person-root">
          ${imageWithPresenceTemplate} ${detailsTemplate}
        </div>
      `;
    }

    if (this.personCardInteraction !== PersonCardInteraction.none) {
      personTemplate = this.renderFlyout(personTemplate, person, image, presence);
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

    const avatarClasses = {
      'avatar-icon': true,
      'ms-Icon': true,
      'ms-Icon--Contact': true,
      small: !this.isLargeAvatar()
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
  protected renderImage(personDetails: IDynamicPerson, imageSrc: string) {
    const title =
      personDetails && this.personCardInteraction === PersonCardInteraction.none
        ? personDetails.displayName || getEmailFromGraphEntity(personDetails) || ''
        : '';

    if (imageSrc && !this._isInvalidImageSrc) {
      return html`
        <div class="img-wrapper">
          <img alt=${title} src=${imageSrc} @error=${() => (this._isInvalidImageSrc = true)} />
        </div>
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
      <div class="user-presence" title=${presence.activity} aria-label=${presence.activity}>
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
  protected renderAvatar(
    personDetails: IDynamicPerson,
    image: string,
    presence: MicrosoftGraphBeta.Presence
  ): TemplateResult {
    const title =
      this.personDetails && this.personCardInteraction === PersonCardInteraction.none
        ? this.personDetails.displayName || getEmailFromGraphEntity(this.personDetails) || ''
        : '';

    const imageClasses = {
      initials: !image || this._isInvalidImageSrc,
      small: !this.isLargeAvatar(),
      'user-avatar': true
    };

    if ((!image || this._isInvalidImageSrc) && this.personDetails) {
      // add avatar background color
      imageClasses[this._personAvatarBg] = true;
    }

    const imageTemplate: TemplateResult = this.renderImage(personDetails, image);
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
    if (!person || (!this.showEmail && !this.showName && this.view === PersonViewType.avatar)) {
      return html``;
    }

    const details: TemplateResult[] = [];

    if (this.showName || this.view > PersonViewType.avatar) {
      const text = this.getTextFromProperty(person, this.line1Property);
      if (text) {
        details.push(html`
          <div class="line1" aria-label="${text}">${text}</div>
        `);
      }
    }

    if (this.showEmail || this.view > PersonViewType.oneline) {
      const text = this.getTextFromProperty(person, this.line2Property);
      if (text) {
        details.push(html`
          <div class="line2" aria-label="${text}">${text}</div>
        `);
      }
    }

    const detailsClasses = classMap({
      details: true,
      small: !this.isLargeAvatar()
    });

    return html`
      <div class="${detailsClasses}">
        ${details}
      </div>
    `;
  }

  /**
   * Render the details flyout.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPerson
   */
  protected renderFlyout(
    anchor: TemplateResult,
    personDetails: IDynamicPerson,
    image: string,
    presence: MicrosoftGraphBeta.Presence
  ): TemplateResult {
    const flyoutContent = this._personCardShouldRender
      ? html`
          <div slot="flyout">
            ${this.renderFlyoutContent(personDetails, image, presence)}
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
  protected renderFlyoutContent(
    personDetails: IDynamicPerson,
    image: string,
    presence: MicrosoftGraphBeta.Presence
  ): TemplateResult {
    return (
      this.renderTemplate('person-card', { person: personDetails, personImage: image }) ||
      html`
        <mgt-person-card
          .personDetails=${personDetails}
          .personImage=${image}
          .personPresence=${presence}
          .showPresence=${this.showPresence}
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
      if (
        !this.personDetails.personImage &&
        ((this.fetchImage && !this.personImage && !this._fetchedImage) || this.personImage === '@')
      ) {
        const image = await getPersonImage(graph, this.personDetails);
        if (image) {
          this.personDetails.personImage = image;
          this._fetchedImage = image;
        }
      }
    } else if (this.userId || this.personQuery === 'me') {
      // Use userId or 'me' query to get the person and image
      const person = await getUserWithPhoto(graph, this.userId);

      this.personDetails = person;
      this._fetchedImage = this.getImage();
    } else if (this.personQuery) {
      // Use the personQuery to find our person.
      const people = await findPeople(graph, this.personQuery, 1);

      if (people && people.length) {
        this.personDetails = people[0];
        const image = await getPersonImage(graph, people[0]);

        if (image) {
          this.personDetails.personImage = image;
          this._fetchedImage = image;
        }
      }
    }

    // populate presence
    const defaultPresence = {
      activity: 'Offline',
      availability: 'Offline',
      id: null
    };
    if (this.showPresence && !this.personPresence && !this._fetchedPresence) {
      try {
        if (this.personDetails && this.personDetails.id) {
          this._fetchedPresence = await getUserPresence(graph, this.personDetails.id);
        } else {
          this._fetchedPresence = defaultPresence;
        }
      } catch (_) {
        // set up a default Presence in case beta api changes or getting error code
        this._fetchedPresence = defaultPresence;
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

    if (this._fetchedImage) {
      return this._fetchedImage;
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

  private getTextFromProperty(personDetails: IDynamicPerson, prop: string) {
    if (!prop || prop.length === 0) {
      return null;
    }

    const properties = prop.trim().split(',');
    let text;
    let i = 0;

    while (!text && i < properties.length) {
      const currentProp = properties[i].trim();
      switch (currentProp) {
        case 'mail':
        case 'email':
          text = getEmailFromGraphEntity(personDetails);
          break;
        default:
          text = personDetails[currentProp];
      }
      i++;
    }

    return text;
  }

  private isLargeAvatar() {
    return (
      this.avatarSize === 'large' ||
      (this.avatarSize === 'auto' && ((this.showEmail && this.showName) || this.view > PersonViewType.oneline))
    );
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
