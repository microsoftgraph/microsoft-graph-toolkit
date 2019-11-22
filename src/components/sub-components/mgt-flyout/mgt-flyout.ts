import { css, customElement, html, LitElement, property } from 'lit-element';
import { PersonCardInteraction } from '../../PersonCardInteraction';

/**
 *
 *
 * @export
 * @class mgt-flyout
 * @extends {MgtTemplatedComponent}
 */
@customElement('mgt-flyout')
export class MgtFlyout extends LitElement {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return css`
      .title {
        color: red;
      }
    `;
  }

  // TODO: create new type
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
  public personCardInteraction: PersonCardInteraction = PersonCardInteraction.hover;

  /**
   *
   *
   * @type {string}
   * @memberof MgtComponent
   */
  @property({
    attribute: 'my-title',
    type: String
  })
  public myTitle: string = 'My First Component';
  private _mouseLeaveTimeout;
  private _mouseEnterTimeout;
  private _openLeft: boolean = false;
  private _openUp: boolean = false;
  @property({ attribute: false }) private _isPersonCardVisible: boolean = false;
  @property({ attribute: false }) private _personCardShouldRender: boolean = false;

  /**
   * Synchronizes property values when attributes change.
   *
   * @param {*} name
   * @param {*} oldValue
   * @param {*} newValue
   * @memberof MgtPersonCard
   */
  public attributeChangedCallback(name, oldval, newval) {
    super.attributeChangedCallback(name, oldval, newval);

    // TODO: handle when an attribute changes.
    //
    // Ex: load data when the name attribute changes
    // if (name === 'person-id' && oldval !== newval){
    //  this.loadData();
    // }
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
  public firstUpdated() {}

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    return html`
      <div
        class="root"
        @mouseenter=${this._handleMouseEnter}
        @mouseleave=${this._handleMouseLeave}
        @click=${this._handleMouseClick}
      >
        <slot></slot>
        <slot name="flyout"></slot>
      </div>
    `;
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

    // TODO expose an event to do this outside of the component
    // const personCard = this.querySelector('mgt-person-card') as MgtPersonCard;
    // if (personCard) {
    //   personCard.isExpanded = false;
    // }
  }
}
