/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Contact, Presence } from '@microsoft/microsoft-graph-types';
import { customElement, html, internalProperty, property, TemplateResult } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { ifDefined } from 'lit-html/directives/if-defined';
import { findPeople, getEmailFromGraphEntity } from '../../graph/graph.people';
import { getGroupImage, getPersonImage } from '../../graph/graph.photos';
import { getUserPresence } from '../../graph/graph.presence';
import { getUserWithPhoto } from '../../graph/graph.userWithPhoto';
import { findUsers, getMe, getUser } from '../../graph/graph.user';
import { AvatarSize, IDynamicPerson, ViewType } from '../../graph/types';
import { Providers, ProviderState, MgtTemplatedComponent } from '@microsoft/mgt-element';
import '../../styles/style-helper';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';
import { MgtPersonCard } from '../mgt-person-card/mgt-person-card';
import '../sub-components/mgt-flyout/mgt-flyout';
import { MgtFlyout } from '../sub-components/mgt-flyout/mgt-flyout';
import { PersonCardInteraction } from './../PersonCardInteraction';
import { styles } from './mgt-person-css';
import { strings } from './strings';

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
  twolines = 4,

  /**
   * Render the avatar and three lines of text
   */
  threelines = 5
}

export enum avatarType {
  /**
   * Renders avatar photo if available, falls back to initials
   */
  photo = 'photo',

  /**
   * Forces render avatar initials
   */
  initials = 'initials'
}

/**
 * Configuration object for the Person component
 *
 * @export
 * @interface MgtPersonConfig
 */
export interface MgtPersonConfig {
  /**
   * Sets or gets whether the person component can use Contacts APIs to
   * find contacts and their images
   *
   * @type {boolean}
   */
  useContactApis: boolean;
}

/**
 * Person properties part of original set provided by graph by default
 */
const defaultPersonProperties = [
  'businessPhones',
  'displayName',
  'givenName',
  'jobTitle',
  'mail',
  'mobilePhone',
  'officeLocation',
  'preferredLanguage',
  'surname',
  'userPrincipalName',
  'id'
];

/**
 * The person component is used to display a person or contact by using their photo, name, and/or email address.
 *
 * @export
 * @class MgtPerson
 * @extends {MgtTemplatedComponent}
 *
 * @fires line1clicked - Fired when line1 is clicked
 * @fires line2clicked - Fired when line2 is clicked
 * @fires line3clicked - Fired when line3 is clicked
 *
 * @cssprop --avatar-size - {Length} Avatar size
 * @cssprop --avatar-border - {String} Avatar border
 * @cssprop --avatar-border-radius - {String} Avatar border radius
 * @cssprop --avatar-cursor - {String} Avatar cursor
 * @cssprop --initials-color - {Color} Initials color
 * @cssprop --initials-background-color - {Color} Initials background color
 * @cssprop --font-family - {String} Font family
 * @cssprop --font-size - {Length} Font size
 * @cssprop --font-weight - {Length} Font weight
 * @cssprop --color - {Color} Color
 * @cssprop --presence-background-color - {Color} Presence badge background color
 * @cssprop --presence-icon-color - {Color} Presence badge icon color
 * @cssprop --text-transform - {String} text transform
 * @cssprop --line2-font-size - {Length} Line 2 font size
 * @cssprop --line2-font-weight - {Length} Line 2 font weight
 * @cssprop --line2-color - {Color} Line 2 color
 * @cssprop --line2-text-transform - {String} Line 2 text transform
 * @cssprop --line3-font-size - {Length} Line 2 font size
 * @cssprop --line3-font-weight - {Length} Line 2 font weight
 * @cssprop --line3-color - {Color} Line 2 color
 * @cssprop --line3-text-transform - {String} Line 2 text transform
 * @cssprop --details-spacing - {Length} spacing between avatar and person details
 */
@customElement('mgt-person')
export class MgtPerson extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  protected get strings() {
    return strings;
  }

  /**
   * Global Configuration object for all
   * person components
   *
   * @static
   * @type {MgtPersonConfig}
   * @memberof MgtPerson
   */
  public static config: MgtPersonConfig = {
    useContactApis: true
  };

  /**
   * allows developer to define name of person for component
   * @type {string}
   */
  @property({
    attribute: 'person-query'
  })
  public get personQuery(): string {
    return this._personQuery;
  }
  public set personQuery(value: string) {
    if (value === this._personQuery) {
      return;
    }

    this._personQuery = value;
    this.personDetailsInternal = null;
    this.requestStateUpdate();
  }

  /**
   * Fallback when no user is found
   * @type {IDynamicPerson}
   */
  @property({
    attribute: 'fallback-details',
    type: Object
  })
  public get fallbackDetails(): IDynamicPerson {
    return this._fallbackDetails;
  }
  public set fallbackDetails(value: IDynamicPerson) {
    if (value === this._fallbackDetails) {
      return;
    }

    this._fallbackDetails = value;

    if (this.personDetailsInternal) {
      return;
    }

    if (value && value.displayName) {
      this._personAvatarBg = this.getColorFromName(value.displayName);
    } else {
      this._personAvatarBg = 'gray20';
    }
    this.requestStateUpdate();
  }

  /**
   * user-id property allows developer to use id value to determine person
   * @type {string}
   */
  @property({
    attribute: 'user-id'
  })
  public get userId(): string {
    return this._userId;
  }
  public set userId(value: string) {
    if (value === this._userId) {
      return;
    }

    this._userId = value;
    this.personDetailsInternal = null;
    this.requestStateUpdate();
  }

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
   * a copy of person-details attribute
   * @type {IDynamicPerson}
   */
  @property({
    attribute: null,
    type: Object
  })
  private get personDetailsInternal(): IDynamicPerson {
    return this._personDetailsInternal;
  }

  private set personDetailsInternal(value: IDynamicPerson) {
    if (this._personDetailsInternal === value) {
      return;
    }

    this._personDetailsInternal = value;
    if (value && value.displayName) {
      this._personAvatarBg = this.getColorFromName(value.displayName);
    } else {
      this._personAvatarBg = 'gray20';
    }

    this._fetchedImage = null;
    this._fetchedPresence = null;

    this.requestStateUpdate();
    this.requestUpdate('personDetailsInternal');
  }

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
   * from the Microsoft Graph based on the personDetailsInternal
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
   * Sets whether to disable the person image fetch
   * from the Microsoft Graph
   *
   * @type {boolean}
   * @memberof MgtPerson
   */
  @property({
    attribute: 'disable-image-fetch',
    type: Boolean
  })
  public disableImageFetch: boolean;

  /**
   * Determines and sets person avatar
   *
   *
   * @type {string}
   * @memberof MgtPerson
   */
  @property({
    attribute: 'avatar-type',
    converter: value => {
      value = value.toLowerCase();

      if (value === 'initials') {
        return avatarType.initials;
      } else {
        return avatarType.photo;
      }
    }
  })
  public get avatarType(): string {
    return this._avatarType;
  }
  public set avatarType(value: string) {
    if (value === this._avatarType) {
      return;
    }

    this._avatarType = value;
    this.requestStateUpdate();
  }

  /**
   * Gets or sets presence of person
   *
   * @type {MicrosoftGraph.Presence}
   * @memberof MgtPerson
   */
  @property({
    attribute: 'person-presence',
    type: Object
  })
  public get personPresence(): Presence {
    return this._personPresence || this._fetchedPresence;
  }
  public set personPresence(value: Presence) {
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
   * Get the scopes required for person
   *
   * @static
   * @return {*}  {string[]}
   * @memberof MgtPerson
   */
  public static get requiredScopes(): string[] {
    const scopes = ['user.readbasic.all', 'user.read', 'people.read', 'presence.read.all', 'presence.read'];

    if (MgtPerson.config.useContactApis) {
      scopes.push('contacts.read');
    }

    return scopes;
  }

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
   * Sets the property of the personDetailsInternal to use for the first line of text.
   * Default is displayName.
   *
   * @type {string}
   * @memberof MgtPerson
   */
  @property({ attribute: 'line1-property' }) public line1Property: string;

  /**
   * Sets the property of the personDetailsInternal to use for the second line of text.
   * Default is mail.
   *
   * @type {string}
   * @memberof MgtPerson
   */
  @property({ attribute: 'line2-property' }) public line2Property: string;

  /**
   * Sets the property of the personDetailsInternal to use for the second line of text.
   * Default is mail.
   *
   * @type {string}
   * @memberof MgtPerson
   */
  @property({ attribute: 'line3-property' }) public line3Property: string;

  /**
   * Sets what data to be rendered (image only, oneLine, twoLines).
   * Default is 'image'.
   *
   * @type {ViewType | PersonViewType}
   * @memberof MgtPerson
   */
  @property({
    converter: value => {
      if (!value || value.length === 0) {
        return ViewType.image;
      }

      value = value.toLowerCase();

      if (typeof ViewType[value] === 'undefined') {
        return ViewType.image;
      } else {
        return ViewType[value];
      }
    }
  })
  public view: ViewType | PersonViewType;

  @internalProperty() private _fetchedImage: string;
  @internalProperty() private _fetchedPresence: Presence;
  @internalProperty() private _isInvalidImageSrc: boolean;
  @internalProperty() private _personCardShouldRender: boolean;

  private _personDetailsInternal: IDynamicPerson;
  private _personDetails: IDynamicPerson;
  private _fallbackDetails: IDynamicPerson;
  private _personAvatarBg: string;
  private _personImage: string;
  private _personPresence: Presence;
  private _personQuery: string;
  private _userId: string;
  private _avatarType: string;

  private _mouseLeaveTimeout;
  private _mouseEnterTimeout;

  constructor() {
    super();

    // defaults
    this.personCardInteraction = PersonCardInteraction.none;
    this.line1Property = 'displayName';
    this.line2Property = 'email';
    this.line3Property = 'jobTitle';
    this.view = ViewType.image;
    this.avatarSize = 'auto';
    this.disableImageFetch = false;
    this._isInvalidImageSrc = false;
    this._avatarType = 'photo';
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  public render() {
    // Loading
    if (this.isLoadingState && !this.personDetails && !this.personDetailsInternal && !this.fallbackDetails) {
      return this.renderLoading();
    }

    // Prep data
    const person = this.personDetails || this.personDetailsInternal || this.fallbackDetails;
    const image = this.getImage();
    const presence = this.personPresence || this._fetchedPresence;

    if (!person && !image) {
      return this.renderNoData();
    }
    if (!(person && person.personImage) && image) {
      person.personImage = image;
    }

    // Default template
    let personTemplate = this.renderTemplate('default', { person, personImage: image, personPresence: presence });

    if (!personTemplate) {
      const detailsTemplate: TemplateResult = this.renderDetails(person, presence);
      const imageWithPresenceTemplate: TemplateResult = this.renderAvatar(person, image, presence);

      const rootClasses = {
        'person-root': true,
        clickable: this.personCardInteraction === PersonCardInteraction.click,
        small: !this.isLargeAvatar()
      };

      personTemplate = html`
        <div 
          class=${classMap(rootClasses)}
          tabindex=${ifDefined(this.personCardInteraction === PersonCardInteraction.click ? '0' : undefined)}
        >
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
        dir=${this.direction}
        @click=${this.handleMouseClick}
        @mouseenter=${this.handleMouseEnter}
        @mouseleave=${this.handleMouseLeave}
        @keydown=${this.handleKeyDown}
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
   * Clears state of the component
   *
   * @protected
   * @memberof MgtPerson
   */
  protected clearState(): void {
    this._personImage = '';
    this._personDetailsInternal = null;
    this._fetchedImage = null;
    this._fetchedPresence = null;
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
   * @param {IDynamicPerson} [personDetailsInternal]
   * @returns
   * @memberof MgtPerson
   */
  protected renderImage(personDetailsInternal: IDynamicPerson, imageSrc: string) {
    if (imageSrc && !this._isInvalidImageSrc && this._avatarType === 'photo') {
      const altText = `${this.strings.photoFor} ${personDetailsInternal.displayName}`;
      return html`
        <div class="img-wrapper">
          <img alt=${altText} src=${imageSrc} @error=${() => (this._isInvalidImageSrc = true)} />
        </div>
      `;
    } else if (personDetailsInternal) {
      const initials = this.getInitials(personDetailsInternal);

      return html`
        <span class="initials-text" aria-label="${this.strings.initials} ${initials}">
          ${
            initials && initials.length
              ? html`
                ${initials}
              `
              : html`
                <i class="ms-Icon ms-Icon--Contact contact-icon"></i>
              `
          }
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
  protected renderPresence(presence: Presence): TemplateResult {
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
          case 'OffWork':
            statusClass = 'presence-offline';
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
      <div class="user-presence" title=${presence.activity} aria-label=${presence.activity} role="img">
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
  protected renderAvatar(personDetailsInternal: IDynamicPerson, image: string, presence: Presence): TemplateResult {
    const hasInitials = !image || this._isInvalidImageSrc || this._avatarType === avatarType.initials;
    const imageClasses = {
      initials: hasInitials,
      small: !this.isLargeAvatar(),
      'user-avatar': true
    };

    let title = '';

    if (hasInitials && personDetailsInternal) {
      // add avatar background color
      imageClasses[this._personAvatarBg] = true;
      title = `${this.strings.initials} ${this.getInitials(personDetailsInternal)}`;
    } else {
      title = personDetailsInternal ? personDetailsInternal.displayName || '' : '';
      if (title !== '') {
        title = `${this.strings.photoFor} ${title}`;
      }
    }

    if (title === '') {
      const emailAddress = getEmailFromGraphEntity(personDetailsInternal);
      if (emailAddress !== null) {
        title = `${this.strings.emailAddress} ${emailAddress}`;
      }
    }

    const imageTemplate: TemplateResult = this.renderImage(personDetailsInternal, image);
    const presenceTemplate: TemplateResult = this.renderPresence(presence);

    return html`
      <div class=${classMap(imageClasses)} title=${title} aria-label=${title}>
        ${imageTemplate} ${presenceTemplate}
      </div>
    `;
  }

  private handleLine1Clicked() {
    this.fireCustomEvent('line1clicked', this.personDetailsInternal);
  }

  private handleLine2Clicked() {
    this.fireCustomEvent('line2clicked', this.personDetailsInternal);
  }

  private handleLine3Clicked() {
    this.fireCustomEvent('line3clicked', this.personDetailsInternal);
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
  protected renderDetails(personProps: IDynamicPerson, presence?: Presence): TemplateResult {
    if (!personProps || this.view === ViewType.image || this.view === PersonViewType.avatar) {
      return html``;
    }

    let person: IDynamicPerson & { presenceActivity?: string; presenceAvailability?: string } = personProps;
    if (presence) {
      person.presenceActivity = presence?.activity;
      person.presenceAvailability = presence?.availability;
    }

    const details: TemplateResult[] = [];

    if (this.view > ViewType.image) {
      if (this.hasTemplate('line1')) {
        // Render the line1 template
        const template = this.renderTemplate('line1', { person });
        details.push(html`
          <div class="line1" @click=${() => this.handleLine1Clicked()}>${template}</div>
        `);
      } else {
        // Render the line1 property value
        const text = this.getTextFromProperty(person, this.line1Property);
        if (text) {
          details.push(html`
            <div class="line1" @click=${() =>
              this.handleLine1Clicked()} role="presentation" aria-label="${text}">${text}</div>
          `);
        }
      }
    }

    if (this.view > ViewType.oneline) {
      if (this.hasTemplate('line2')) {
        // Render the line2 template
        const template = this.renderTemplate('line2', { person });
        details.push(html`
          <div class="line2" @click=${() => this.handleLine2Clicked()}>${template}</div>
        `);
      } else {
        // Render the line2 property value
        const text = this.getTextFromProperty(person, this.line2Property);
        if (text) {
          details.push(html`
            <div class="line2" @click=${() =>
              this.handleLine2Clicked()} role="presentation" aria-label="${text}">${text}</div>
          `);
        }
      }
    }

    if (this.view > ViewType.twolines) {
      if (this.hasTemplate('line3')) {
        // Render the line3 template
        const template = this.renderTemplate('line3', { person });
        details.push(html`
          <div class="line3" @click=${() => this.handleLine3Clicked()}>${template}</div>
        `);
      } else {
        // Render the line3 property value
        const text = this.getTextFromProperty(person, this.line3Property);
        if (text) {
          details.push(html`
            <div class="line3" @click=${() =>
              this.handleLine3Clicked()} role="presentation" aria-label="${text}">${text}</div>
          `);
        }
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
    presence: Presence
  ): TemplateResult {
    const flyoutContent = this._personCardShouldRender
      ? html`
          <div slot="flyout">
            ${this.renderFlyoutContent(personDetails, image, presence)}
          </div>
        `
      : html``;

    return html`
      <mgt-flyout light-dismiss class="flyout" .avoidHidingAnchor=${false}>
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
  protected renderFlyoutContent(personDetails: IDynamicPerson, image: string, presence: Presence): TemplateResult {
    return (
      this.renderTemplate('person-card', { person: personDetails, personImage: image }) ||
      html`
        <mgt-person-card
          lock-tab-navigation
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

    if (provider && provider.state === ProviderState.SignedOut) {
      this.personDetailsInternal = null;
      return;
    }

    const graph = provider.graph.forComponent(this);

    // Prepare person props
    let personProps = [...defaultPersonProperties, this.line1Property, this.line2Property, this.line3Property];
    personProps = personProps.filter(email => email !== 'email');

    let details = this.personDetailsInternal || this.personDetails;

    if (details) {
      if (
        !details.personImage &&
        this.fetchImage &&
        this._avatarType === 'photo' &&
        !this.personImage &&
        !this._fetchedImage
      ) {
        let image: string;
        if ('groupTypes' in details) {
          image = await getGroupImage(graph, details, MgtPerson.config.useContactApis);
        } else {
          image = await getPersonImage(graph, details, MgtPerson.config.useContactApis);
        }
        if (image) {
          details.personImage = image;
          this._fetchedImage = image;
        }
      }
    } else if (this.userId || this.personQuery === 'me') {
      // Use userId or 'me' query to get the person and image
      let person;
      if (this._avatarType === 'photo' && !this.disableImageFetch) {
        person = await getUserWithPhoto(graph, this.userId, personProps);
      } else {
        if (this.personQuery === 'me') {
          person = await getMe(graph, personProps);
        } else {
          person = await getUser(graph, this.userId, personProps);
        }
      }
      this.personDetailsInternal = person;
      this.personDetails = person;
      this._fetchedImage = this.getImage();
    } else if (this.personQuery) {
      // Use the personQuery to find our person.
      let people = await findPeople(graph, this.personQuery, 1);

      if (!people || people.length === 0) {
        people = (await findUsers(graph, this.personQuery, 1)) || [];
      }

      if (people && people.length) {
        this.personDetailsInternal = people[0];
        this.personDetails = people[0];
        if (this._avatarType === 'photo' && !this.disableImageFetch) {
          const image = await getPersonImage(graph, people[0], MgtPerson.config.useContactApis);

          if (image) {
            this.personDetailsInternal.personImage = image;
            this.personDetails.personImage = image;
            this._fetchedImage = image;
          }
        }
      }
    }

    // populate presence
    const defaultPresence: Presence = {
      activity: 'Offline',
      availability: 'Offline',
      id: null
    };

    if (this.showPresence && !this.personPresence && !this._fetchedPresence) {
      try {
        details = this.personDetailsInternal || this.personDetails;
        if (details) {
          // setting userId to 'me' ensures only the presence.read permission is required
          const userId = this.personQuery !== 'me' ? details?.id : null;
          this._fetchedPresence = await getUserPresence(graph, userId);
        } else {
          this._fetchedPresence = defaultPresence;
        }
      } catch (_) {
        // set up a default Presence in case beta api changes or getting error code
        this._fetchedPresence = defaultPresence;
      }
    }
  }

  /**
   * Gets the user initials
   *
   * @protected
   * @returns {string}
   * @memberof MgtPerson
   */
  protected getInitials(person?: IDynamicPerson): string {
    if (!person) {
      person = this.personDetailsInternal;
    }

    if ((person as Contact).initials) {
      return (person as Contact).initials;
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

  /**
   * Gets color from name
   *
   * @protected
   * @param {string} name
   * @returns {string}
   * @memberof MgtPerson
   */
  protected getColorFromName(name: string): string {
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

  private getImage(): string {
    if (this.personImage) {
      return this.personImage;
    }

    if (this._fetchedImage) {
      return this._fetchedImage;
    }

    const person = this.personDetailsInternal || this.personDetails;
    return person && person.personImage ? person.personImage : null;
  }

  private isLetter(char: string) {
    try {
      return char.match(new RegExp('\\p{L}', 'u'));
    } catch (e) {
      return char.toLowerCase() !== char.toUpperCase();
    }
  }

  private getTextFromProperty(personDetailsInternal: IDynamicPerson, prop: string) {
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
          text = getEmailFromGraphEntity(personDetailsInternal);
          break;
        default:
          text = personDetailsInternal[currentProp];
      }
      i++;
    }

    return text;
  }

  private isLargeAvatar() {
    return this.avatarSize === 'large' || (this.avatarSize === 'auto' && this.view > ViewType.oneline);
  }

  private handleMouseClick(e: MouseEvent) {
    if (this.personCardInteraction === PersonCardInteraction.click) {
      this.showPersonCard();
    }
  }

  private handleKeyDown(e: KeyboardEvent) {
    //enter activates person-card
    if (e) {
      if (e.key === 'Enter') {
        this.showPersonCard();
      }
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
      flyout.close();
    }
    const personCard = (this.querySelector('mgt-person-card') ||
      this.renderRoot.querySelector('mgt-person-card')) as MgtPersonCard;
    if (personCard) {
      personCard.isExpanded = false;
      personCard.clearHistory();
    }
  }

  private showPersonCard() {
    if (!this._personCardShouldRender) {
      this._personCardShouldRender = true;
    }

    const flyout = this.flyout;
    if (flyout) {
      flyout.open();
    }
  }
}
