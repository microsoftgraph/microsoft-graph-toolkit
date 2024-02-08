/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/* eslint-disable @typescript-eslint/no-unused-expressions */
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
import { registerMgtThemeToggleComponent } from './mgt-theme-toggle';

import { html, fixture, expect } from '@open-wc/testing';
import { sendMouse, emulateMedia } from '@web/test-runner-commands';
class Deferred<T = unknown> {
  promise: Promise<T>;
  resolve: (value: T) => void;
  reject: (reason?: unknown) => void;
  constructor() {
    this.promise = new Promise((resolve, reject) => {
      this.resolve = resolve;
      this.reject = reject;
    });
  }
}

const getMiddleOfElement = (element: Element) => {
  const { x, y, width, height } = element.getBoundingClientRect();

  return {
    x: Math.floor(x + window.scrollX + width / 2),
    y: Math.floor(y + window.scrollY + height / 2)
  };
};

describe('mgt-theme-toggle - media behavior tests', () => {
  registerMgtThemeToggleComponent();

  it('should render as checked whe color scheme is dark', async () => {
    await emulateMedia({ colorScheme: 'dark' });
    expect(matchMedia('(prefers-color-scheme: dark)').matches).to.be.true;
    const darkToggle = await fixture('<mgt-theme-toggle></mgt-theme-toggle>');
    await expect(darkToggle).shadowDom.to.equal(
      `<fluent-switch
        aria-checked="true"
        aria-disabled="false"
        checked="true"
        role="switch"
        tabindex="0"
      >
        <span slot="checked-message">Dark</span>
        <span slot="unchecked-message">Light</span>
        <label for="direction-switch">Color mode:</label>
      </fluent-switch>`
    );
  });
  it('should emit darkmodechanged true on inital render when the color scheme is dark', async () => {
    const deferred = new Deferred<boolean>();
    const listener = (e: CustomEvent<boolean>) => {
      deferred.resolve(e.detail);
    };
    await emulateMedia({ colorScheme: 'dark' });
    expect(matchMedia('(prefers-color-scheme: dark)').matches).to.be.true;
    await fixture(html`<mgt-theme-toggle @darkmodechanged=${listener}></mgt-theme-toggle>`);
    // darkmodechanged emitted when setting initial value of checked
    expect(await deferred.promise).to.be.true;
  });

  it('should render as unchecked when color scheme is light', async () => {
    await emulateMedia({ colorScheme: 'light' });
    expect(matchMedia('(prefers-color-scheme: light)').matches).to.be.true;
    const lightToggle = await fixture('<mgt-theme-toggle></mgt-theme-toggle>');
    await expect(lightToggle).shadowDom.to.equal(
      `<fluent-switch
        aria-checked="false"
        aria-disabled="false"
        checked="false"
        role="switch"
        tabindex="0"
      >
        <span slot="checked-message">Dark</span>
        <span slot="unchecked-message">Light</span>
        <label for="direction-switch">Color mode:</label>
      </fluent-switch>`
    );
  });
});

describe('mgt-theme-toggle - tests', () => {
  beforeEach(() => {
    registerMgtThemeToggleComponent();
  });
  it('should render', async () => {
    const toggle = await fixture('<mgt-theme-toggle></mgt-theme-toggle>');
    await expect(toggle).shadowDom.to.equal(
      `<fluent-switch
        aria-checked="false"
        aria-disabled="false"
        checked="false"
        role="switch"
        tabindex="0"
      >
        <span slot="checked-message">Dark</span>
        <span slot="unchecked-message">Light</span>
        <label for="direction-switch">Color mode:</label>
      </fluent-switch>`
    );
  });

  it("should emit darkmodechanged with the current 'checked' state on click", async () => {
    let darkModeDeferred = new Deferred<boolean>();
    const listener = (e: CustomEvent<boolean>) => {
      darkModeDeferred.resolve(e.detail);
    };

    const element = await fixture(html`<mgt-theme-toggle @darkmodechanged=${listener}></mgt-theme-toggle>`);
    const toggle: HTMLInputElement = element.shadowRoot.querySelector('[role=switch]');
    const { x, y } = getMiddleOfElement(element);

    // darkmodechanged emitted when setting initial value of checked
    expect(await darkModeDeferred.promise).to.be.false;
    expect(toggle).attribute('checked', 'false');

    // click to set switch on
    darkModeDeferred = new Deferred<boolean>();
    await sendMouse({ type: 'click', position: [x, y] });

    expect(await darkModeDeferred.promise).to.be.true;
    expect(toggle.checked).to.be.true;

    // click to set switch off
    darkModeDeferred = new Deferred<boolean>();
    await sendMouse({ type: 'click', position: [x, y] });

    expect(await darkModeDeferred.promise).to.be.false;
    expect(toggle.checked).to.be.false;
  });

  it('should have a checked switch if mode is dark', async () => {
    const toggle = await fixture('<mgt-theme-toggle mode="dark"></mgt-theme-toggle>');
    await expect(toggle).shadowDom.to.equal(
      `<fluent-switch
        aria-checked="true"
        aria-disabled="false"
        checked="true"
        role="switch"
        tabindex="0"
      >
        <span slot="checked-message">Dark</span>
        <span slot="unchecked-message">Light</span>
        <label for="direction-switch">Color mode:</label>
      </fluent-switch>`
    );
    await expect(toggle).shadowDom.to.be.accessible();
  });
});
