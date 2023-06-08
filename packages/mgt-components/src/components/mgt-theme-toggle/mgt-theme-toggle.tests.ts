/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
// import the mock for media match first to ensure it's hoisted and available for our dependencies
import '../../__mocks__/mock-media-match';
import { screen } from 'testing-library__dom';
import { fixture } from '@open-wc/testing-helpers';
import './mgt-theme-toggle';

class Deferred<T = unknown> {
  promise: Promise<T>;
  resolve: (value: T) => void;
  reject: (reason?: any) => void;
  constructor() {
    this.promise = new Promise((resolve, reject) => {
      this.resolve = resolve;
      this.reject = reject;
    });
  }
}

describe('mgt-theme-toggle - tests', () => {
  it('should render', async () => {
    await fixture('<mgt-theme-toggle></mgt-theme-toggle>');
    const toggle = await screen.findByRole('switch');
    expect(toggle).not.toBeNull();
  });
  it("should emit darkmodechanged with the current 'checked' state on click", async () => {
    let darkModeState = false;
    const element = await fixture('<mgt-theme-toggle></mgt-theme-toggle>');
    const toggle: HTMLInputElement = await screen.findByRole('switch');
    expect(toggle).not.toBeNull();
    let deferred = new Deferred<boolean>();
    element.addEventListener('darkmodechanged', (e: CustomEvent<boolean>) => {
      deferred.resolve(e.detail);
    });
    expect(darkModeState).toBe(false);
    expect(toggle.checked).toBe(false);
    toggle.click();
    expect(toggle.checked).toBe(true);
    darkModeState = await deferred.promise;
    expect(darkModeState).toBe(true);
    deferred = new Deferred<boolean>();
    toggle.click();
    expect(toggle.checked).toBe(false);
    darkModeState = await deferred.promise;
    expect(darkModeState).toBe(false);
  });

  it('should have a checked switch if mode is dark', async () => {
    await fixture('<mgt-theme-toggle mode="dark"></mgt-theme-toggle>');
    const toggle = await screen.findByRole('switch');
    expect(toggle).not.toBeNull();
    expect(toggle.getAttribute('aria-checked')).toBe('true');
    expect(toggle.getAttribute('checked')).toBe('true');
  });

  it('should not have a checked switch if mode is light', async () => {
    await fixture('<mgt-theme-toggle></mgt-theme-toggle>');
    const toggle = await screen.findByRole('switch');
    expect(toggle).not.toBeNull();
    expect(toggle.getAttribute('aria-checked')).toBe('false');
    expect(toggle.getAttribute('checked')).toBe('false');
  });

  it('should have a checked switch if user prefers dark mode and no mode is set', async () => {
    // redefine matchMedia to return true
    Object.defineProperty(window, 'matchMedia', {
      writable: true,
      value: jest.fn().mockImplementation((query: unknown) => ({
        matches: true,
        media: query,
        onchange: null,
        addListener: jest.fn(), // deprecated
        removeListener: jest.fn(), // deprecated
        addEventListener: jest.fn(),
        removeEventListener: jest.fn(),
        dispatchEvent: jest.fn()
      }))
    });
    await fixture('<mgt-theme-toggle></mgt-theme-toggle>');
    const toggle = await screen.findByRole('switch');
    expect(toggle).not.toBeNull();
    expect(toggle.getAttribute('aria-checked')).toBe('true');
    expect(toggle.getAttribute('checked')).toBe('true');
  });
});
