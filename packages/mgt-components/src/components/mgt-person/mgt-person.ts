/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { MgtTemplatedComponent, ProviderState, Providers, customElement, mgtHtml } from '@microsoft/mgt-element';
import { Contact, Presence } from '@microsoft/microsoft-graph-types';
import { html, TemplateResult } from 'lit';
import { property, state } from 'lit/decorators.js';
import { classMap } from 'lit/directives/class-map.js';
import { findPeople, getEmailFromGraphEntity } from '../../graph/graph.people';
import { getGroupImage, getPersonImage } from '../../graph/graph.photos';
import { getUserPresence } from '../../graph/graph.presence';
import { findUsers, getMe, getUser } from '../../graph/graph.user';
import { getUserWithPhoto } from '../../graph/graph.userWithPhoto';
import { AvatarSize, IDynamicPerson, ViewType } from '../../graph/types';
import '../../styles/style-helper';
import { SvgIcon, getSvg } from '../../utils/SvgHelper';
import { MgtPersonCard } from '../mgt-person-card/mgt-person-card';
import '../sub-components/mgt-flyout/mgt-flyout';
import { MgtFlyout } from '../sub-components/mgt-flyout/mgt-flyout';
import { PersonCardInteraction } from './../PersonCardInteraction';
import { styles } from './mgt-person-css';
import { MgtPersonConfig, PersonViewType, avatarType } from './mgt-person-types';
import { strings } from './strings';

export { PersonCardInteraction } from '../PersonCardInteraction';

/**
 * Person properties part of original set provided by graph by default
 */
const defaultPersonProperties = [
  'businessPhones',
  'displayName',
  'givenName',
  'jobTitle',
  'department',
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
 * @fires {CustomEvent<IDynamicPerson>} line1clicked - Fired when line1 is clicked
 * @fires {CustomEvent<IDynamicPerson>} line2clicked - Fired when line2 is clicked
 * @fires {CustomEvent<IDynamicPerson>} line3clicked - Fired when line3 is clicked
 * @fires {CustomEvent<IDynamicPerson>} line4clicked - Fired when line4 is clicked
 *
 * @cssprop --person-background-color - {Color} the color of the person component background.
 * @cssprop --person-background-border-radius - {Length} the border radius of the person component. Default is 4px.
 *
 * @cssprop --person-avatar-size-small - {Length} the width and height of the avatar. Default is 24px.
 * @cssprop --person-avatar-size-2-lines - {Length} the width and height of the avatar. Default is 40px.
 * @cssprop --person-avatar-size-3-lines - {Length} the width and height of the avatar. Default is 56px.
 * @cssprop --person-avatar-size-4-lines - {Length} the width and height of the avatar. Default is 72px.
 * @cssprop --person-avatar-size-vertical - {Length} the width and height of the avatar. Default is 72px.
 * @cssprop --person-avatar-border - {String} the border around an avatar. Default is none.
 * @cssprop --person-avatar-border-radius - {String} the radius around the border of an avatar. Default is 50%.
 *
 * @cssprop --person-initials-text-color - {Color} the color of initials in an avatar.
 * @cssprop --person-initials-background-color - {Color} the color of the background in an avatar with initials.
 *
 * @cssprop --person-details-spacing - {Length} the space between the avatar and the person details. Default is 12px.
 *
 * @cssprop --person-line1-font-size - {String} the font-size of the line 1 text. Default is 14px.
 * @cssprop --person-line1-font-weight - {Length} the font weight of the line 1 text. Default is 600.
 * @cssprop --person-line1-text-color - {Color} the color of the line 1 text.
 * @cssprop --person-line1-text-transform - {String} the tex transform of the line 1 text. Default is inherit.
 * @cssprop --person-line1-text-line-height - {Length} the line height of the line 1 text. Default is 20px.
 *
 * @cssprop --person-line2-font-size - {Length} the font-size of the line 2 text. Default is 12px.
 * @cssprop --person-line2-font-weight - {Length} the font weight of the line 2 text. Default is 400.
 * @cssprop --person-line2-text-color - {Color} the color of the line 2 text.
 * @cssprop --person-line2-text-transform - {String} the tex transform of the line 2 text. Default is inherit.
 * @cssprop --person-line2-text-line-height - {Length} the line height of the line 2 text. Default is 16px.
 *
 * @cssprop --person-line3-font-size - {Length} the font-size of the line 3 text. Default is 12px.
 * @cssprop --person-line3-font-weight - {Length} the font weight of the line 3 text. Default is 400.
 * @cssprop --person-line3-text-color - {Color} the color of the line 3 text.
 * @cssprop --person-line3-text-transform - {String} the tex transform of the line 3 text. Default is inherit.
 * @cssprop --person-line3-text-line-height - {Length} the line height of the line 3 text. Default is 16px.
 *
 * @cssprop --person-line4-font-size - {Length} the font-size of the line 4 text. Default is 12px.
 * @cssprop --person-line4-font-weight - {Length} the font weight of the line 4 text. Default is 400.
 * @cssprop --person-line4-text-color - {Color} the color of the line 4 text.
 * @cssprop --person-line4-text-transform - {String} the tex transform of the line 4 text. Default is inherit.
 * @cssprop --person-line4-text-line-height - {Length} the line height of the line 4 text. Default is 16px.
 *
 * @cssprop --person-details-wrapper-width - {Length} the minimum width of the details section. Default is 168px.
 */
@customElement('person')
export class MgtPerson extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * Strings to use for localization
   *
   * @readonly
   * @protected
   * @memberof MgtPerson
   */
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
   *
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
    void this.requestStateUpdate();
  }

  /**
   * Fallback when no user is found
   *
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

    void this.requestStateUpdate();
  }

  /**
   * user-id property allows developer to use id value to determine person
   *
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
    void this.requestStateUpdate();
  }

  /**
   * usage property allows you to specify where the component is being used to add
   * customized personalization for it. Currently only supports "people" as used in
   * the people component.
   *
   * @type {string}
   */
  @property({
    attribute: 'usage'
  })
  public get usage(): string {
    return this._usage;
  }
  public set usage(value: string) {
    if (value === this._usage) {
      return;
    }

    this._usage = value;
    void this.requestStateUpdate();
  }

  /**
   * determines if person component renders presence
   *
   * @type {boolean}
   */
  @property({
    attribute: 'show-presence',
    type: Boolean
  })
  public showPresence: boolean;

  /**
   * determines person component avatar size and apply presence badge accordingly.
   * Default is "auto". When you set the view > 1, it will default to "auto".
   *
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
   *
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
    this._fetchedImage = null;
    this._fetchedPresence = null;

    void this.requestStateUpdate();
    this.requestUpdate('personDetailsInternal');
  }

  /**
   * object containing Graph details on person
   *
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
    this._fetchedImage = null;
    this._fetchedPresence = null;

    void this.requestStateUpdate();
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
   * Sets the vertical layout of
   * the Person Card
   *
   * @type {boolean}
   * @memberof MgtPerson
   */
  @property({
    attribute: 'vertical-layout',
    type: Boolean
  })
  public verticalLayout: boolean;

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
    void this.requestStateUpdate();
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
    converter: value => {
      value = value.toLowerCase();
      if (typeof PersonCardInteraction[value] === 'undefined') {
        return PersonCardInteraction.none;
      } else {
        return PersonCardInteraction[value] as PersonCardInteraction;
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
   * Sets the property of the personDetailsInternal to use for the third line of text.
   * Default is mail.
   *
   * @type {string}
   * @memberof MgtPerson
   */
  @property({ attribute: 'line3-property' }) public line3Property: string;

  /**
   * Sets the property of the personDetailsInternal to use for the fourth line of text.
   * Default is mail.
   *
   * @type {string}
   * @memberof MgtPerson
   */
  @property({ attribute: 'line4-property' }) public line4Property: string;

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
        return ViewType[value] as ViewType;
      }
    }
  })
  public view: ViewType | PersonViewType;

  @state() private _fetchedImage: string;
  @state() private _fetchedPresence: Presence;
  @state() private _isInvalidImageSrc: boolean;
  @state() private _personCardShouldRender: boolean;

  private _personDetailsInternal: IDynamicPerson;
  private _personDetails: IDynamicPerson;
  private _fallbackDetails: IDynamicPerson;
  private _personImage: string;
  private _personPresence: Presence;
  private _personQuery: string;
  private _userId: string;
  private _usage: string;
  private _avatarType: string;

  private _mouseLeaveTimeout = -1;
  private _mouseEnterTimeout = -1;

  constructor() {
    super();

    // defaults
    this.personCardInteraction = PersonCardInteraction.none;
    this.line1Property = 'displayName';
    this.line2Property = 'jobTitle';
    this.line3Property = 'department';
    this.line4Property = 'email';
    this.view = ViewType.image;
    this.avatarSize = 'auto';
    this.disableImageFetch = false;
    this._isInvalidImageSrc = false;
    this._avatarType = 'photo';
    this.verticalLayout = false;
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

      personTemplate = html`
        ${imageWithPresenceTemplate}
        ${detailsTemplate}
      `;
    }

    const showPersonCard = this.personCardInteraction !== PersonCardInteraction.none;
    if (showPersonCard) {
      personTemplate = this.renderFlyout(personTemplate, person, image, presence);
    }

    const rootClasses = classMap({
      'person-root': true,
      small: !this.isThreeLines() && !this.isFourLines() && !this.isLargeAvatar(),
      large: this.avatarSize !== 'auto' && this.isLargeAvatar(),
      noline: this.isNoLine(),
      oneline: this.isOneLine(),
      twolines: this.isTwoLines(),
      threelines: this.isThreeLines(),
      fourlines: this.isFourLines(),
      vertical: this.isVertical()
    });

    return html`
      <div
        class="${rootClasses}"
        dir=${this.direction}
        @click=${this.handleMouseClick}
        @mouseenter=${this.handleMouseEnter}
        @mouseleave=${this.handleMouseLeave}
        @keydown=${this.handleKeyDown}>
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
      vertical: this.isVertical(),
      small: !this.isLargeAvatar(),
      threeLines: this.isThreeLines(),
      fourLines: this.isFourLines()
    };

    return html`
       <i class=${classMap(avatarClasses)}></i>
     `;
  }

  /**
   * Render a person icon.
   *
   * @protected
   * @returns
   * @memberof MgtPerson
   */
  protected renderPersonIcon() {
    return getSvg(SvgIcon.Person);
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
    const altText = `${this.strings.photoFor} ${personDetailsInternal.displayName}`;
    const hasImage = imageSrc && !this._isInvalidImageSrc && this._avatarType === 'photo';
    const imageTemplate = html`<img alt=${altText} src=${imageSrc} @error=${() => (this._isInvalidImageSrc = true)} />`;

    const initials = personDetailsInternal ? this.getInitials(personDetailsInternal) : '';
    const hasInitials = initials?.length;
    const textClasses = classMap({
      initials: hasInitials && !hasImage,
      'contact-icon': !hasInitials
    });
    const contactIconTemplate = html`<i>${this.renderPersonIcon()}</i>`;
    const textTemplate = html`<span class="${textClasses}">${hasInitials ? initials : contactIconTemplate}</span>`;

    return hasImage ? imageTemplate : textTemplate;
  }

  /**
   * Render presence for the person.
   *
   * @param presence
   * @memberof MgtPerson
   * @returns
   */
  protected renderPresence(presence: Presence): TemplateResult {
    if (!this.showPresence || !presence) {
      return html``;
    }
    let presenceIcon: TemplateResult;

    const { activity, availability } = presence;
    switch (availability) {
      case 'Available':
        switch (activity) {
          case 'Available':
            presenceIcon = getSvg(SvgIcon.PresenceAvailable);
            break;
          case 'OutOfOffice':
            presenceIcon = getSvg(SvgIcon.PresenceOofAvailable);
            break;
        }
        break;
      case 'Busy':
        switch (activity) {
          case 'Busy':
          case 'InACall':
          case 'InAMeeting':
            presenceIcon = getSvg(SvgIcon.PresenceBusy);
            break;
          case 'OutOfOffice':
          case 'OnACall':
            presenceIcon = getSvg(SvgIcon.PresenceOofBusy);
            break;
        }
        break;
      case 'DoNotDisturb':
        switch (activity) {
          case 'OutOfOffice':
            presenceIcon = getSvg(SvgIcon.PresenceOofDnd);
            break;
          default:
            presenceIcon = getSvg(SvgIcon.PresenceDnd);
            break;
        }
        break;

      case 'Away':
        switch (activity) {
          case 'OutOfOffice':
            presenceIcon = getSvg(SvgIcon.PresenceOofAway);
            break;
          default:
            presenceIcon = getSvg(SvgIcon.PresenceAway);
            break;
        }
        break;
      case 'Offline':
        switch (activity) {
          case 'Offline':
            presenceIcon = getSvg(SvgIcon.PresenceOffline);
            break;
          case 'OutOfOffice':
            presenceIcon = getSvg(SvgIcon.PresenceOofAway);
            break;
          default:
            presenceIcon = getSvg(SvgIcon.PresenceStatusUnknown);
            break;
        }
        break;
      default:
        presenceIcon = getSvg(SvgIcon.PresenceStatusUnknown);
        break;
    }

    const presenceWrapperClasses = classMap({
      'presence-wrapper': true,
      noline: this.isNoLine(),
      oneline: this.isOneLine()
    });

    return html`
      <span
        class="${presenceWrapperClasses}"
        title="${availability}"
        aria-label="${availability}"
        role="img">
          ${presenceIcon}
      </span>
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

    let title = '';

    if (hasInitials && personDetailsInternal) {
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
      <div class="avatar-wrapper">
        ${imageTemplate}
        ${presenceTemplate}
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

  private handleLine4Clicked() {
    this.fireCustomEvent('line4clicked', this.personDetailsInternal);
  }

  /**
   * Render the details part of the person template.
   *
   * @param personProps
   * @param presence
   * @memberof MgtPerson
   * @returns
   */
  protected renderDetails(personProps: IDynamicPerson, presence?: Presence): TemplateResult {
    if (!personProps || this.view === ViewType.image || this.view === PersonViewType.avatar) {
      return html``;
    }

    const person: IDynamicPerson & { presenceActivity?: string; presenceAvailability?: string } = personProps;
    if (presence) {
      person.presenceActivity = presence?.activity;
      person.presenceAvailability = presence?.availability;
    }

    const details: TemplateResult[] = [];

    if (this.view > ViewType.image) {
      const text = this.getTextFromProperty(person, this.line1Property);
      if (this.hasTemplate('line1')) {
        // Render the line1 template
        const template = this.renderTemplate('line1', { person });
        details.push(html`
           <div class="line1" @click=${() =>
             this.handleLine1Clicked()} role="presentation" aria-label="${text}">${template}</div>
         `);
      } else {
        // Render the line1 property value
        if (text) {
          details.push(html`
             <div class="line1" @click=${() =>
               this.handleLine1Clicked()} role="presentation" aria-label="${text}">${text}</div>
           `);
        }
      }
    }

    if (this.view > ViewType.oneline) {
      const text = this.getTextFromProperty(person, this.line2Property);
      if (this.hasTemplate('line2')) {
        // Render the line2 template
        const template = this.renderTemplate('line2', { person });
        details.push(html`
           <div class="line2" @click=${() =>
             this.handleLine2Clicked()} role="presentation" aria-label="${text}">${template}</div>
         `);
      } else {
        // Render the line2 property value
        if (text) {
          details.push(html`
             <div class="line2" @click=${() =>
               this.handleLine2Clicked()} role="presentation" aria-label="${text}">${text}</div>
           `);
        }
      }
    }

    if (this.view > ViewType.twolines) {
      const text = this.getTextFromProperty(person, this.line3Property);
      if (this.hasTemplate('line3')) {
        // Render the line3 template
        const template = this.renderTemplate('line3', { person });
        details.push(html`
           <div class="line3" @click=${() =>
             this.handleLine3Clicked()} role="presentation" aria-label="${text}">${template}</div>
         `);
      } else {
        // Render the line3 property value
        if (text) {
          details.push(html`
             <div class="line3" @click=${() =>
               this.handleLine3Clicked()} role="presentation" aria-label="${text}">${text}</div>
           `);
        }
      }
    }

    if (this.view > ViewType.threelines) {
      const text = this.getTextFromProperty(person, this.line4Property);
      if (this.hasTemplate('line4')) {
        // Render the line4 template
        const template = this.renderTemplate('line4', { person });
        details.push(html`
          <div class="line4" @click=${() =>
            this.handleLine4Clicked()} role="presentation" aria-label="${text}">${template}</div>
        `);
      } else {
        // Render the line4 property value
        if (text) {
          details.push(html`
            <div class="line4" @click=${() =>
              this.handleLine4Clicked()} role="presentation" aria-label="${text}">${text}</div>
          `);
        }
      }
    }

    const detailsClasses = classMap({
      'details-wrapper': true,
      vertical: this.isVertical()
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
           <div slot="flyout" data-testid="flyout-slot">
             ${this.renderFlyoutContent(personDetails, image, presence)}
           </div>`
      : html``;

    const slotClasses = classMap({
      vertical: this.isVertical()
    });

    return mgtHtml`
      <mgt-flyout light-dismiss class="flyout" .avoidHidingAnchor=${false}>
        <div slot="anchor" class="${slotClasses}">${anchor}</div>
        ${flyoutContent}
      </mgt-flyout>`;
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
      mgtHtml`
        <mgt-person-card
          lock-tab-navigation
          .personDetails=${personDetails}
          .personImage=${image}
          .personPresence=${presence}
          .showPresence=${this.showPresence}>
        </mgt-person-card>`
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

    if (this.fallbackDetails) {
      this.line2Property = 'email';
    }

    if (this.verticalLayout && this.view < ViewType.fourlines) {
      this.line2Property = 'email';
    }

    // Prepare person props
    let personProps = [
      ...defaultPersonProperties,
      this.line1Property,
      this.line2Property,
      this.line3Property,
      this.line4Property
    ];
    personProps = personProps.filter(email => email !== 'email');

    let details = this.personDetailsInternal || this.personDetails || this.fallbackDetails;

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
      let person: IDynamicPerson;
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

  private getTextFromProperty(personDetailsInternal: IDynamicPerson, prop: string): string {
    if (!prop || prop.length === 0) {
      return null;
    }

    const properties = prop.trim().split(',');
    let text: string;
    let i = 0;

    while (!text && i < properties.length) {
      const currentProp = properties[i].trim();
      switch (currentProp) {
        case 'mail':
        case 'email':
          text = getEmailFromGraphEntity(personDetailsInternal);
          break;
        default:
          text = personDetailsInternal[currentProp] as string;
      }
      i++;
    }

    return text;
  }

  private isLargeAvatar() {
    return this.avatarSize === 'large' || (this.avatarSize === 'auto' && this.view > ViewType.oneline);
  }

  private isNoLine() {
    return this.view < ViewType.oneline;
  }

  private isOneLine() {
    return this.view === ViewType.oneline;
  }

  private isTwoLines() {
    return this.view === ViewType.twolines;
  }

  private isThreeLines() {
    return this.view === ViewType.threelines;
  }

  private isFourLines() {
    return this.view === ViewType.fourlines;
  }

  private isVertical() {
    return this.verticalLayout;
  }

  private handleMouseClick = (e: MouseEvent) => {
    const element = e.target as HTMLElement;
    if (this.personCardInteraction === PersonCardInteraction.click && element.tagName !== 'MGT-PERSON-CARD') {
      this.showPersonCard();
    }
  };

  private handleKeyDown = (e: KeyboardEvent) => {
    // enter activates person-card
    if (e) {
      if (e.key === 'Enter') {
        this.showPersonCard();
      }
    }
  };

  private handleMouseEnter = (e: MouseEvent) => {
    clearTimeout(this._mouseEnterTimeout);
    clearTimeout(this._mouseLeaveTimeout);
    if (this.personCardInteraction !== PersonCardInteraction.hover) {
      return;
    }
    this._mouseEnterTimeout = window.setTimeout(this.showPersonCard, 500);
  };

  private handleMouseLeave = (e: MouseEvent) => {
    clearTimeout(this._mouseEnterTimeout);
    clearTimeout(this._mouseLeaveTimeout);
    this._mouseLeaveTimeout = window.setTimeout(this.hidePersonCard, 500);
  };

  /**
   * hides the person card
   *
   * @memberof MgtPerson
   */
  public hidePersonCard = () => {
    const flyout = this.flyout;
    if (flyout) {
      flyout.close();
    }
    const personCard =
      this.querySelector<MgtPersonCard>('mgt-person-card') || this.renderRoot.querySelector('mgt-person-card');
    if (personCard) {
      personCard.isExpanded = false;
      personCard.clearHistory();
    }
  };

  public showPersonCard = () => {
    if (!this._personCardShouldRender) {
      this._personCardShouldRender = true;
    }

    const flyout = this.flyout;
    if (flyout) {
      flyout.open();
    }
  };
}
