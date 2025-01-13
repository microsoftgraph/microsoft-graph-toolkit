/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html, TemplateResult } from 'lit';
import { property } from 'lit/decorators.js';
import { MgtBaseTaskComponent } from '@microsoft/mgt-element';
import { fluentSwitch } from '@fluentui/web-components';
import { registerFluentComponents } from '../../utils/FluentComponents';
import { applyTheme } from '../../styles/theme-manager';
import { strings } from './strings';
import { registerComponent } from '@microsoft/mgt-element';

export const registerMgtThemeToggleComponent = () => {
  registerFluentComponents(fluentSwitch);
  registerComponent('theme-toggle', MgtThemeToggle);
};

/**
 * Toggle to switch between light and dark mode
 * Will detect browser preference and set accordingly or dark mode can be forced
 *
 * @fires {CustomEvent<boolean>} darkmodechanged - Fired when dark mode is toggled by a user action
 *
 * @class MgtDarkToggle
 * @extends {MgtBaseTaskComponent}
 *
 * @fires {CustomEvent<undefined>} updated - Fired when the component is updated
 */
export class MgtThemeToggle extends MgtBaseTaskComponent {
  constructor() {
    super();
    const prefersDarkMode = window.matchMedia('(prefers-color-scheme:dark)').matches;
    this.darkModeActive = prefersDarkMode;
    this.applyTheme(this.darkModeActive);
  }
  /**
   * Provides strings for localization
   *
   * @readonly
   * @protected
   * @memberof MgtDarkToggle
   */
  protected get strings() {
    return strings;
  }

  /**
   * Controls whether dark mode is active
   *
   * @type {boolean}
   * @memberof MgtDarkToggle
   */
  @property({
    attribute: 'mode',
    reflect: true,
    type: String,
    converter: {
      fromAttribute: (value: string) => {
        return value === 'dark';
      },
      toAttribute: (value: boolean) => {
        return value ? 'dark' : 'light';
      }
    }
  })
  public darkModeActive: boolean;

  /**
   * Fires after a component is updated.
   * Allows a component to trigger side effects after updating.
   *
   * @param {Map<string, any>} changedProperties
   * @memberof MgtDarkToggle
   */
  updated(changedProperties: Map<string, unknown>): void {
    if (changedProperties.has('darkModeActive')) {
      this.applyTheme(this.darkModeActive);
    }
  }

  /**
   * renders the component
   *
   * @return {TemplateResult}
   * @memberof MgtDarkToggle
   */
  render(): TemplateResult {
    return html`
      <fluent-switch checked=${this.darkModeActive} @change=${this.onSwitchChanged}>
        <span slot="checked-message">${strings.on}</span>
        <span slot="unchecked-message">${strings.off}</span>
        <label for="direction-switch">${strings.label}</label>
      </fluent-switch>
`;
  }

  private readonly onSwitchChanged = (e: Event) => {
    this.darkModeActive = (e.target as HTMLInputElement).checked;
  };

  private applyTheme(active: boolean) {
    const targetTheme = active ? 'dark' : 'light';
    applyTheme(targetTheme);

    document.body.classList.remove('mgt-dark-mode', 'mgt-light-mode');
    document.body.classList.add(`mgt-${targetTheme}-mode`);
    this.fireCustomEvent('darkmodechanged', this.darkModeActive, true, false, true);
  }
}
