/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { AuthenticationProvider } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProvider';
import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';
import { Graph } from '../Graph';
import { EventDispatcher } from './EventDispatcher';
/**
 * Provider Type to be extended for implmenting new providers
 *
 * @export
 * @abstract
 * @class IProvider
 * @implements {AuthenticationProvider}
 */
export abstract class IProvider implements AuthenticationProvider {
  // tslint:disable-next-line: completed-docs
  public graph: Graph;
  private _state: ProviderState;
  private _loginChangedDispatcher = new EventDispatcher<LoginChangedEvent>();
  /**
   * returns state of Provider
   *
   * @readonly
   * @type {ProviderState}
   * @memberof IProvider
   */
  public get state(): ProviderState {
    return this._state;
  }

  constructor() {
    this._state = ProviderState.Loading;
  }

  /**
   * sets state of Provider and fires loginchangedDispatcher
   *
   * @param {ProviderState} state
   * @memberof IProvider
   */
  public setState(state: ProviderState) {
    if (state !== this._state) {
      this._state = state;
      this._loginChangedDispatcher.fire({});
    }
  }

  /**
   * event handler when login changes
   *
   * @param {EventHandler<LoginChangedEvent>} eventHandler
   * @memberof IProvider
   */
  public onStateChanged(eventHandler: EventHandler<LoginChangedEvent>) {
    this._loginChangedDispatcher.add(eventHandler);
  }
  /**
   * removes event andler for when login changes
   *
   * @param {EventHandler<LoginChangedEvent>} eventHandler
   * @memberof IProvider
   */
  public removeStateChangedHandler(eventHandler: EventHandler<LoginChangedEvent>) {
    this._loginChangedDispatcher.remove(eventHandler);
  }

  /**
   * option implementation that can be called to sign in user (required for mgt-login to work)
   *
   * @returns {Promise<void>}
   * @memberof IProvider
   */
  public login?(): Promise<void>;

  /**
   * optional implementation that can be called to sign out user (required for mgt-login to work)
   *
   * @returns {Promise<void>}
   * @memberof IProvider
   */
  public logout?(): Promise<void>;

  /**
   * uses scopes to recieve access token
   *
   * @param {...string[]} scopes
   * @returns {Promise<string>}
   * @memberof IProvider
   */
  public getAccessTokenForScopes(...scopes: string[]): Promise<string> {
    return this.getAccessToken({ scopes });
  }

  /**
   * Promise to receive access token using Provider options
   *
   * @abstract
   * @param {AuthenticationProviderOptions} [options]
   * @returns {Promise<string>}
   * @memberof IProvider
   */
  public abstract getAccessToken(options?: AuthenticationProviderOptions): Promise<string>;
  // tslint:disable-next-line: completed-docs
}

// tslint:disable-next-line: completed-docs
export type EventHandler<E> = (event: E) => void;

/**
 * loginChangedEvent
 *
 * @export
 * @interface LoginChangedEvent
 */
// tslint:disable-next-line: no-empty-interface
export interface LoginChangedEvent {}

/**
 * LoginType
 *
 * @export
 * @enum {number}
 */
export enum LoginType {
  /**
   * Popup = 0
   */
  Popup,
  /**
   * Redirect = 1
   */
  Redirect
}

/**
 * ProviderState
 *
 * @export
 * @enum {number}
 */
export enum ProviderState {
  /**
   * Loading = 0
   */
  Loading,
  /**
   * SignedOut = 1
   */
  SignedOut,
  /**
   * SignedIn = 1
   */
  SignedIn
}
