import { css, customElement, html, LitElement, property, PropertyValues } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { styles } from './mgt-flyout-css';
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
    return styles;
  }

  /**
   *
   *
   * @type {string}
   * @memberof MgtComponent
   */
  @property({
    attribute: 'isOpen',
    type: Boolean
  })
  public isOpen: boolean = false;

  private renderedOnce = false;

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

    const anchor = this.renderRoot.querySelector('.anchor');
    const flyout = this.renderRoot.querySelector('.flyout') as HTMLElement;
    if (flyout && anchor) {
      const flyoutRect = flyout.getBoundingClientRect();
      const anchorRect = anchor.getBoundingClientRect();

      const windowWidth = window.innerWidth || document.documentElement.clientWidth;
      const windowHeight = window.innerHeight || document.documentElement.clientHeight;

      let left: number;
      let right: number;
      let top: number;
      let bottom: number;

      if (anchorRect.width >= flyoutRect.width) {
        left = 0;
      } else {
        const centerOffset = flyoutRect.width / 2 - anchorRect.width / 2;

        if (anchorRect.left - centerOffset < 0) {
          left = -anchorRect.left;
        } else {
          left = -centerOffset;
        } // todo check if offscreen to the right and move left
      }

      flyout.style.left = left ? `${left}px` : '';
      flyout.style.right = right ? `${right}px` : '';
      flyout.style.top = top ? `${top}px` : '';
      flyout.style.bottom = bottom ? `${bottom}px` : '';
    }
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    return html`
      <div class="root">
        <div class="anchor">
          <slot></slot>
        </div>
        ${this.renderFlyout()}
      </div>
    `;
  }

  private renderFlyout() {
    if (!this.isOpen && !this.renderedOnce) {
      return;
    }

    this.renderedOnce = true;

    const classes = {
      flyout: true,
      visible: this.isOpen
    };

    return html`
      <div class=${classMap(classes)}>
        <slot name="flyout"></slot>
      </div>
    `;
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
