/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { fluentSkeleton } from '@fluentui/web-components';
import {
  MgtTemplatedTaskComponent,
  ProviderState,
  Providers,
  buildComponentName,
  customElementHelper,
  mgtHtml,
  registerComponent
} from '@microsoft/mgt-element';
import { Presence } from '@microsoft/microsoft-graph-types';
import { TemplateResult, html, nothing } from 'lit';
import { property, state } from 'lit/decorators.js';
import { classMap } from 'lit/directives/class-map.js';
import { ifDefined } from 'lit/directives/if-defined.js';
import { isContact, isUser } from '../../graph/entityType';
import { findPeople, getEmailFromGraphEntity } from '../../graph/graph.people';
import { getGroupImage, getPersonImage } from '../../graph/graph.photos';
import { getUserPresence } from '../../graph/graph.presence';
import { findUsers, getMe, getUser } from '../../graph/graph.user';
import { getUserWithPhoto } from '../../graph/graph.userWithPhoto';
import { AvatarSize, IDynamicPerson, ViewType, viewTypeConverter } from '../../graph/types';
import '../../styles/style-helper';
import { registerFluentComponents } from '../../utils/FluentComponents';
import { SvgIcon, getSvg } from '../../utils/SvgHelper';
import { debounce } from '../../utils/Utils';
import { IExpandable, IHistoryClearer } from '../mgt-person-card/types';
import '../sub-components/mgt-flyout/mgt-flyout';
import { MgtFlyout, registerMgtFlyoutComponent } from '../sub-components/mgt-flyout/mgt-flyout';
import { personCardConverter, type PersonCardInteraction } from './../PersonCardInteraction';
import { styles } from './mgt-person-css';
import { AvatarType, MgtPersonConfig, avatarTypeConverter } from './mgt-person-types';
import { strings } from './strings';
import { getPersonCardGraphData } from '../mgt-person-card/mgt-person-card.graph';

/**
 * Person properties part of original set provided by graph by default
 */
export const defaultPersonProperties = [
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
  'id',
  'userType'
];

export const registerMgtPersonComponent = () => {
  registerFluentComponents(fluentSkeleton);

  // register self first to avoid infinte loop due to circular ref between person and person card
  registerComponent('person', MgtPerson);

  registerMgtFlyoutComponent();
};

/**
 * The person component is used to display a person or contact by using their photo, name, and/or email address.
 *
 * @export
 * @class MgtPerson
 * @extends {MgtTemplatedComponent}
 *
 * @fires {CustomEvent<undefined>} updated - Fired when the component is updated
 * @fires {CustomEvent<IDynamicPerson>} line1clicked - Fired when line1 is clicked
 * @fires {CustomEvent<IDynamicPerson>} line2clicked - Fired when line2 is clicked
 * @fires {CustomEvent<IDynamicPerson>} line3clicked - Fired when line3 is clicked
 * @fires {CustomEvent<IDynamicPerson>} line4clicked - Fired when line4 is clicked
 *
 * @cssprop --person-background-color - {Color} the color of the person component background.
 * @cssprop --person-background-border-radius - {Length} the border radius of the person component. Default is 4px.
 *
 * @cssprop --person-avatar-size - {Length} the width and height of the avatar. Default is 24px.
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
export class MgtPerson extends MgtTemplatedTaskComponent {
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
  public get personQuery(): string {
    return this._personQuery;
  }
  @property({
    attribute: 'person-query'
  })
  public set personQuery(value: string) {
    if (value === this._personQuery) {
      return;
    }

    this._personQuery = value;
    this._personDetails = null;
    this.personDetailsInternal = null;
  }

  /**
   * Fallback when no user is found
   *
   * @type {IDynamicPerson}
   */
  public get fallbackDetails(): IDynamicPerson {
    return this._fallbackDetails;
  }
  @property({
    attribute: 'fallback-details',
    type: Object
  })
  public set fallbackDetails(value: IDynamicPerson) {
    if (value === this._fallbackDetails) {
      return;
    }

    this._fallbackDetails = value;

    if (this.personDetailsInternal) {
      return;
    }
  }

  /**
   * user-id property allows developer to use id value to determine person
   *
   * @type {string}
   */
  public get userId(): string {
    return this._userId;
  }
  @property({
    attribute: 'user-id'
  })
  public set userId(value: string) {
    if (value === this._userId) {
      return;
    }

    this._userId = value;
    this.personDetailsInternal = null;
    this._personDetails = null;
  }

  /**
   * usage property allows you to specify where the component is being used to add
   * customized personalization for it. Currently only supports "people" and "people-picker" as used in
   * the people component.
   *
   * @type {string}
   */
  public get usage(): string {
    return this._usage;
  }
  @property({
    attribute: 'usage'
  })
  public set usage(value: string) {
    if (value === this._usage) {
      return;
    }

    this._usage = value;
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
   * Valid options are 'small', 'large', and 'auto'
   * Default is "auto". When you set the view more than oneline, it will default to "auto".
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
  private get personDetailsInternal(): IDynamicPerson {
    return this._personDetailsInternal;
  }

  @state()
  private set personDetailsInternal(value: IDynamicPerson) {
    if (this._personDetailsInternal === value) {
      return;
    }

    this._personDetailsInternal = value;
    this._fetchedImage = null;
    this._fetchedPresence = null;
  }

  /**
   * object containing Graph details on person
   *
   * @type {IDynamicPerson}
   */
  public get personDetails(): IDynamicPerson {
    return this._personDetails;
  }

  @property({
    attribute: 'person-details',
    type: Object
  })
  public set personDetails(value: IDynamicPerson) {
    if (this._personDetails === value) {
      return;
    }

    this._personDetails = value;
    this._fetchedImage = null;
    this._fetchedPresence = null;
  }

  /**
   * Set the image of the person
   *
   * @type {string}
   * @memberof MgtPersonCard
   */
  public get personImage(): string {
    return this._personImage || this._fetchedImage;
  }
  @property({
    attribute: 'person-image',
    type: String
  })
  public set personImage(value: string) {
    if (value === this._personImage) {
      return;
    }

    this._isInvalidImageSrc = !value;
    this._personImage = value;
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
   * Determines and sets person avatar view
   * Valid options are 'photo' or 'initials'
   *
   * @type {AvatarType}
   * @memberof MgtPerson
   */
  @property({
    attribute: 'avatar-type',
    converter: (value): AvatarType => avatarTypeConverter(value, 'photo')
  })
  public avatarType: AvatarType = 'photo';

  /**
   * Gets or sets presence of person
   *
   * @type {MicrosoftGraph.Presence}
   * @memberof MgtPerson
   */
  public get personPresence(): Presence {
    return this._personPresence || this._fetchedPresence;
  }
  @property({
    attribute: 'person-presence',
    type: Object
  })
  public set personPresence(value: Presence) {
    if (value === this._personPresence) {
      return;
    }
    this._personPresence = value;
  }

  /**
   * Sets how the person-card is invoked
   * Valid options are: 'none', 'hover', or 'click'
   * Set to 'none' to not show the card
   *
   * @type {PersonCardInteraction}
   * @memberof MgtPerson
   */
  @property({
    attribute: 'person-card',
    converter: value => personCardConverter(value)
  })
  public personCardInteraction: PersonCardInteraction = 'none';

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
   * Sets what data to be rendered.
   * Valid options are 'image', 'oneline', 'twolines', 'threelines', or 'fourlines'
   * Default is 'image'.
   *
   * @type {ViewType}
   * @memberof MgtPerson
   */
  @property({
    converter: value => viewTypeConverter(value, 'image')
  })
  public view: ViewType = 'image';

  @state() private _fetchedImage: string;
  @state() private _fetchedPresence: Presence;
  @state() private _isInvalidImageSrc: boolean;
  @state() private _personCardShouldRender: boolean;
  @state() private _hasLoadedPersonCard = false;

  private _personDetailsInternal: IDynamicPerson;
  private _personDetails: IDynamicPerson;
  private _fallbackDetails: IDynamicPerson;
  private _personImage: string;
  private _personPresence: Presence;
  private _personQuery: string;
  private _userId: string;
  private _usage: string;

  private _mouseLeaveTimeout = -1;
  private _mouseEnterTimeout = -1;
  private _keyBoardFocus: { (): void; (): void };

  constructor() {
    super();

    // defaults
    this.personCardInteraction = 'none';
    this.line1Property = 'displayName';
    this.line2Property = 'jobTitle';
    this.line3Property = 'department';
    this.line4Property = 'email';
    this.view = 'image';
    this.avatarSize = 'auto';
    this.disableImageFetch = false;
    this._isInvalidImageSrc = false;
    this.avatarType = 'photo';
    this.verticalLayout = false;
  }

  protected readonly renderContent = () => {
    // Prep data
    const person = this.personDetails || this.personDetailsInternal || this.fallbackDetails;
    const image = this.getImage();
    const presence = this.personPresence || this._fetchedPresence;

    if (!person && !image) {
      return this.renderNoData();
    }
    if (!person?.personImage && image) {
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

    const showPersonCard = this.personCardInteraction !== 'none';
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
        @keydown=${this.handleKeyDown}
        tabindex="${ifDefined(this.personCardInteraction !== 'none' ? '0' : undefined)}"
      >
        ${personTemplate}
      </div>
    `;
  };

  /**
   * Render the loading state
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPerson
   */
  protected renderLoading = (): TemplateResult => {
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

    const detailsClasses = classMap({
      'details-wrapper': true,
      vertical: this.isVertical()
    });

    return (
      this.renderTemplate('loading', null) ||
      html`
        <div class="${rootClasses}">
          <div class="avatar-wrapper">
            <fluent-skeleton shimmer class="shimmer" shape="circle"></fluent-skeleton>
          </div>
          <div class=${detailsClasses}>
            ${this.renderLoadingLines()}
          </div>
        </div>`
    );
  };

  protected renderLoadingLines = (): TemplateResult[] => {
    const lines: TemplateResult[] = [];
    if (this.isNoLine()) return lines;
    if (this.isOneLine()) {
      lines.push(this.renderLoadingLine(1));
    }
    if (this.isTwoLines()) {
      lines.push(this.renderLoadingLine(1));
      lines.push(this.renderLoadingLine(2));
    }
    if (this.isThreeLines()) {
      lines.push(this.renderLoadingLine(1));
      lines.push(this.renderLoadingLine(2));
      lines.push(this.renderLoadingLine(3));
    }
    if (this.isFourLines()) {
      lines.push(this.renderLoadingLine(1));
      lines.push(this.renderLoadingLine(2));
      lines.push(this.renderLoadingLine(3));
      lines.push(this.renderLoadingLine(4));
    }
    return lines;
  };

  protected renderLoadingLine = (line: number): TemplateResult => {
    const lineNumber = `line${line}`;
    return html`
      <div class=${lineNumber}>
        <fluent-skeleton shimmer class="shimmer text" shape="rect"></fluent-skeleton>
      </div>
    `;
  };

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
      noline: this.isNoLine(),
      oneline: this.isOneLine(),
      twolines: this.isTwoLines(),
      threelines: this.isThreeLines(),
      fourlines: this.isFourLines()
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
    const hasImage = imageSrc && !this._isInvalidImageSrc && this.avatarType === 'photo';
    const imageOnly = this.avatarType === 'photo' && this.view === 'image';
    const titleText =
      (personDetailsInternal?.displayName ||
        `${this.strings.emailAddress} ${getEmailFromGraphEntity(personDetailsInternal)}`) ??
      undefined;
    const imageTemplate = html`<img
      title="${ifDefined(imageOnly ? titleText : undefined)}"
      alt=${altText}
      src=${imageSrc}
      @error=${() => (this._isInvalidImageSrc = true)} />`;

    const initials = personDetailsInternal ? this.getInitials(personDetailsInternal) : '';
    const hasInitials = initials?.length;
    const textClasses = classMap({
      initials: hasInitials && !hasImage,
      'contact-icon': !hasInitials
    });
    const contactIconTemplate = html`<i>${this.renderPersonIcon()}</i>`;
    // consider the image to presentational if the view is anything other than image.
    // this reduces the redundant announcement of the user's name.
    const textTemplate = html`
      <span 
        title="${ifDefined(this.view === 'image' ? titleText : undefined)}"
        role="${ifDefined(this.view === 'image' ? undefined : 'presentation')}"
        class="${textClasses}"
      >
        ${hasInitials ? initials : contactIconTemplate}
      </span>
`;
    if (hasImage) {
      this.fireCustomEvent('person-image-rendered');
    } else {
      this.fireCustomEvent('person-icon-rendered');
    }

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
      case 'AvailableIdle':
        switch (activity) {
          case 'OutOfOffice':
            presenceIcon = getSvg(SvgIcon.PresenceOofAvailable);
            break;
          // OutOfOffice and Uknowns
          case 'Available':
          default:
            presenceIcon = getSvg(SvgIcon.PresenceAvailable);
            break;
        }
        break;
      case 'Busy':
      case 'BusyIdle':
        switch (activity) {
          case 'OutOfOffice':
          case 'OnACall':
            presenceIcon = getSvg(SvgIcon.PresenceOofBusy);
            break;
          // Busy,InACall,InAConferenceCall,InAMeeting, Unknown
          case 'Busy':
          case 'InACall':
          case 'InAMeeting':
          case 'InAConferenceCall':
          default:
            presenceIcon = getSvg(SvgIcon.PresenceBusy);
            break;
        }
        break;
      case 'DoNotDisturb':
        switch (activity) {
          case 'OutOfOffice':
            presenceIcon = getSvg(SvgIcon.PresenceOofDnd);
            break;
          case 'Presenting':
          case 'Focusing':
          case 'UrgentInterruptionsOnly':
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
          case 'AwayLastSeenTime':
          default:
            presenceIcon = getSvg(SvgIcon.PresenceAway);
            break;
        }
        break;
      case 'BeRightBack':
        switch (activity) {
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
          case 'OffWork':
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

    const formattedActivity = (this.strings[activity] as string) ?? nothing;

    return html`
      <span
        class="${presenceWrapperClasses}"
        title="${formattedActivity}"
        aria-label="${formattedActivity}"
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
    if (!personProps || this.view === 'image') {
      return html``;
    }

    const person: IDynamicPerson & { presenceActivity?: string; presenceAvailability?: string } = personProps;
    if (presence) {
      person.presenceActivity = presence?.activity;
      person.presenceAvailability = presence?.availability;
    }

    const details: TemplateResult[] = [];

    // we already returned on image, so we must have a first line
    const line1text = this.getTextFromProperty(person, this.line1Property);
    if (this.hasTemplate('line1')) {
      // Render the line1 template
      const template = this.renderTemplate('line1', { person });
      details.push(html`
           <div class="line1" part="detail-line" @click=${() =>
             this.handleLine1Clicked()} role="presentation" aria-label="${line1text}">${template}</div>
         `);
    } else {
      // Render the line1 property value
      if (line1text) {
        details.push(html`
             <div class="line1" part="detail-line" @click=${() =>
               this.handleLine1Clicked()} role="presentation" aria-label="${line1text}">${line1text}</div>
           `);
      }
    }

    // if we have more than one line we add the second line
    if (!this.isOneLine()) {
      const text = this.getTextFromProperty(person, this.line2Property);
      if (this.hasTemplate('line2')) {
        // Render the line2 template
        const template = this.renderTemplate('line2', { person });
        details.push(html`
           <div class="line2" part="detail-line" @click=${() =>
             this.handleLine2Clicked()} role="presentation" aria-label="${text}">${template}</div>
         `);
      } else {
        // Render the line2 property value
        if (text) {
          details.push(html`
             <div class="line2" part="detail-line" @click=${() =>
               this.handleLine2Clicked()} role="presentation" aria-label="${text}">${text}</div>
           `);
        }
      }
    }

    // if we have a third or fourth line we add the third line
    if (this.isThreeLines() || this.isFourLines()) {
      const text = this.getTextFromProperty(person, this.line3Property);
      if (this.hasTemplate('line3')) {
        // Render the line3 template
        const template = this.renderTemplate('line3', { person });
        details.push(html`
           <div class="line3" part="detail-line" @click=${() =>
             this.handleLine3Clicked()} role="presentation" aria-label="${text}">${template}</div>
         `);
      } else {
        // Render the line3 property value
        if (text) {
          details.push(html`
             <div class="line3" part="detail-line" @click=${() =>
               this.handleLine3Clicked()} role="presentation" aria-label="${text}">${text}</div>
           `);
        }
      }
    }

    // add the fourth line if necessary
    if (this.isFourLines()) {
      const text = this.getTextFromProperty(person, this.line4Property);
      if (this.hasTemplate('line4')) {
        // Render the line4 template
        const template = this.renderTemplate('line4', { person });
        details.push(html`
          <div class="line4" part="detail-line" @click=${() =>
            this.handleLine4Clicked()} role="presentation" aria-label="${text}">${template}</div>
        `);
      } else {
        // Render the line4 property value
        if (text) {
          details.push(html`
            <div class="line4" part="detail-line" @click=${() =>
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
    const flyoutContent =
      this._personCardShouldRender && this._hasLoadedPersonCard
        ? html`
           <div slot="flyout" data-testid="flyout-slot">
             ${this.renderFlyoutContent(personDetails, image, presence)}
           </div>`
        : html``;

    const slotClasses = classMap({
      small: !this.isThreeLines() && !this.isFourLines() && !this.isLargeAvatar(),
      large: this.avatarSize !== 'auto' && this.isLargeAvatar(),
      noline: this.isNoLine(),
      oneline: this.isOneLine(),
      twolines: this.isTwoLines(),
      threelines: this.isThreeLines(),
      fourlines: this.isFourLines(),
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
    this.fireCustomEvent('flyout-content-rendered');
    return (
      this.renderTemplate('person-card', { person: personDetails, personImage: image }) ||
      mgtHtml`
        <mgt-person-card
          class="mgt-person-card"
          lock-tab-navigation
          .personDetails=${personDetails}
          .personImage=${image}
          .personPresence=${presence}
          .showPresence=${this.showPresence}>
        </mgt-person-card>`
    );
  }

  protected args() {
    return [
      this.providerState,
      this.verticalLayout,
      this.view,
      this.fallbackDetails,
      this.line1Property,
      this.line2Property,
      this.line3Property,
      this.line4Property,
      this.fetchImage,
      this.avatarType,
      this.userId,
      this.personQuery,
      this.disableImageFetch,
      this.showPresence,
      this.personPresence,
      this.personDetails
    ];
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

    if ((this.verticalLayout && this.view !== 'fourlines') || this.fallbackDetails) {
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

    let details = this.personDetailsInternal || this.personDetails;
    if (details) {
      if (
        !details.personImage &&
        this.fetchImage &&
        this.avatarType === 'photo' &&
        !this.personImage &&
        !this._fetchedImage
      ) {
        let image: string;
        if ('groupTypes' in details) {
          image = await getGroupImage(graph, details);
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
      if (this.avatarType === 'photo' && !this.disableImageFetch) {
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

      if (people?.length) {
        this.personDetailsInternal = people[0];
        this.personDetails = people[0];
        if (this.avatarType === 'photo' && !this.disableImageFetch) {
          const image = await getPersonImage(graph, people[0], MgtPerson.config.useContactApis);

          if (image) {
            this.personDetailsInternal.personImage = image;
            this.personDetails.personImage = image;
            this._fetchedImage = image;
          }
        }
      }
    }

    details = this.personDetailsInternal || this.personDetails || this.fallbackDetails;

    // load card data at this point
    if (this.personCardInteraction !== 'none') {
      // perform the batch requests and cache
      void getPersonCardGraphData(graph, details, this.personQuery === 'me');
    }

    // populate presence
    const defaultPresence: Presence = {
      activity: 'Offline',
      availability: 'Offline',
      id: null
    };

    if (this.showPresence && !this.personPresence && !this._fetchedPresence) {
      try {
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

    if (isContact(person)) {
      return person.initials;
    }

    let initials = '';
    if (isUser(person)) {
      initials += person.givenName?.[0]?.toUpperCase() ?? '';
      initials += person.surname?.[0]?.toUpperCase() ?? '';
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
    return person?.personImage ? person.personImage : null;
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
    return (
      this.avatarSize === 'large' || (this.avatarSize === 'auto' && this.view !== 'image' && this.view !== 'oneline')
    );
  }

  private isNoLine() {
    return this.view === 'image';
  }

  private isOneLine() {
    return this.view === 'oneline';
  }

  private isTwoLines() {
    return this.view === 'twolines';
  }

  private isThreeLines() {
    return this.view === 'threelines';
  }

  private isFourLines() {
    return this.view === 'fourlines';
  }

  private isVertical() {
    return this.verticalLayout;
  }

  private readonly handleMouseClick = (e: MouseEvent) => {
    const element = e.target as HTMLElement;
    // todo: fix for disambiguation
    if (
      this.personCardInteraction === 'click' &&
      element.tagName !== `${customElementHelper.prefix}-PERSON-CARD`.toUpperCase()
    ) {
      this.showPersonCard();
    }
  };

  private readonly handleKeyDown = (e: KeyboardEvent) => {
    const personEl = this.renderRoot.querySelector<HTMLElement>('.person-root');
    // enter activates and focuses on person-card
    if (e) {
      if (e.key === 'Enter') {
        this.showPersonCard();
        const flyout = this.flyout;
        if (flyout?.isOpen) {
          this._keyBoardFocus = debounce(() => {
            const personCardEl = flyout.querySelector<HTMLElement>('.mgt-person-card');
            personCardEl.setAttribute('tabindex', '0');
            personCardEl.focus();
          }, 500);
          this._keyBoardFocus();
        }
        personEl.blur();
      }
      if (this.personCardInteraction !== 'none') {
        if (e.key === 'Escape' && personEl) {
          this.hidePersonCard();
          personEl.focus();
        }
      }
    }
  };

  private readonly handleMouseEnter = () => {
    clearTimeout(this._mouseEnterTimeout);
    clearTimeout(this._mouseLeaveTimeout);
    if (this.personCardInteraction !== 'hover') {
      return;
    }
    this._mouseEnterTimeout = window.setTimeout(this.showPersonCard, 500);
  };

  private readonly handleMouseLeave = () => {
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
      this.querySelector<Element & IExpandable & IHistoryClearer>('.mgt-person-card') ||
      this.renderRoot.querySelector('.mgt-person-card');
    if (personCard) {
      personCard.isExpanded = false;
      personCard.clearHistory();
    }
  };

  private readonly loadPersonCardResources = async () => {
    // if there could be a person-card then we should load those resources using a dynamic import
    if (this.personCardInteraction !== 'none' && !this._hasLoadedPersonCard) {
      const { registerMgtPersonCardComponent } = await import('../mgt-person-card/mgt-person-card');

      // only register person card if it hasn't been registered yet
      if (!customElements.get(buildComponentName('person-card'))) registerMgtPersonCardComponent();

      this._hasLoadedPersonCard = true;
    }
  };

  public showPersonCard = () => {
    if (!this._personCardShouldRender) {
      this._personCardShouldRender = true;
      void this.loadPersonCardResources();
    }

    const flyout = this.flyout;
    if (flyout) {
      flyout.open();
    }
  };
}
