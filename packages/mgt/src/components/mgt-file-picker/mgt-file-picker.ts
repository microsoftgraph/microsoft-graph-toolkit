import { MgtTemplatedComponent } from '@microsoft/mgt-element';
import { customElement, html, TemplateResult } from 'lit-element';
import { MgtFlyout } from '../sub-components/mgt-flyout/mgt-flyout';

import { styles } from './mgt-file-picker-css';

@customElement('mgt-file-picker')
export class MgtFilePicker extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
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

  render() {
    const root = html`
      <div class="root">
        ${this.renderButton()}
      </div>
    `;

    const flyoutContent = html`
      <div slot="flyout">
        ${this.renderFlyoutContent()}
      </div>
    `;

    return html`
      <mgt-flyout light-dismiss class="flyout">
        ${root} ${flyoutContent}
      </mgt-flyout>
    `;
  }

  protected renderButton(): TemplateResult {
    return html`
      <div class="button" @click=${() => this.openFlyout()}>
        Choose from OneDrive
      </div>
    `;
  }

  protected renderFlyoutContent(): TemplateResult {
    return html`
      Files
    `;
  }

  protected openFlyout(): void {
    this.flyout.open();
  }
}
