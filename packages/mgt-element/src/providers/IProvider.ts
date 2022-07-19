/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { AuthenticationProvider } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProvider';
import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';
import { IGraph } from '../IGraph';
import { EventDispatcher, EventHandler } from '../utils/EventDispatcher';

/**
 * Provider Type to be extended for implmenting new providers
 *
 * @export
 * @abstract
 * @class IProvider
 * @implements {AuthenticationProvider}
 */
export abstract class IProvider implements AuthenticationProvider {
  /**
   * The Graph object that contains the Graph client sdk
   *
   * @type {Graph}
   * @memberof IProvider
   */
  public graph: IGraph;
  /**
   * Enable/Disable multi account functionality
   *
   * @protected
   * @type {boolean}
   * @memberof IProvider
   */
  protected isMultipleAccountDisabled: boolean = true;
  private _state: ProviderState;
  private _loginChangedDispatcher = new EventDispatcher<LoginChangedEvent>();
  private _activeAccountChangedDispatcher = new EventDispatcher<ActiveAccountChanged>();

  /**
   * Enable/Disable incremental consent
   *
   * @protected
   * @type {boolean}
   * @memberof IProvider
   */
  private _isIncrementalConsentDisabled: boolean = false;

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

  /**
   * Incremental consent setting
   *
   * @readonly
   * @memberof IProvider
   */
  public get isIncrementalConsentDisabled(): boolean {
    return this._isIncrementalConsentDisabled;
  }

  /**
   * Enable/Disable incremental consent
   *
   * @readonly
   * @memberof IProvider
   */
  public set isIncrementalConsentDisabled(disabled: boolean) {
    this._isIncrementalConsentDisabled = disabled;
  }

  /**
   * Name used for analytics
   *
   * @readonly
   * @memberof IProvider
   */
  public get name() {
    return 'MgtIProvider';
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
   * removes event handler for when login changes
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
   * Returns all signed in accounts.
   *
   * @return {*}  {any[]}
   * @memberof IProvider
   */
  public getAllAccounts?(): IProviderAccount[];

  /**
   * Switch between two signed in accounts
   *
   * @param {*} user
   * @memberof IProvider
   */
  public setActiveAccount?(user: IProviderAccount) {
    this.fireActiveAccountChanged();
  }

  /**
   * Event handler when Active account changes
   *
   * @param {EventHandler<ActiveAccountChanged>} eventHandler
   * @memberof IProvider
   */
  public onActiveAccountChanged(eventHandler: EventHandler<ActiveAccountChanged>) {
    this._activeAccountChangedDispatcher.add(eventHandler);
  }

  /**
   * Removes event handler for when Active account changes
   *
   * @param {EventHandler<ActiveAccountChanged>} eventHandler
   * @memberof IProvider
   */
  public removeActiveAccountChangedHandler(eventHandler: EventHandler<ActiveAccountChanged>) {
    this._activeAccountChangedDispatcher.remove(eventHandler);
  }

  /**
   * Fires event when active account changes
   *
   * @memberof IProvider
   */
  private fireActiveAccountChanged() {
    this._activeAccountChangedDispatcher.fire({});
  }

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
}

/**
 * ActiveAccountChanged Event
 *
 * @export
 * @interface ActiveAccountChanged
 */
export interface ActiveAccountChanged {}
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
   * SignedIn = 2
   */
  SignedIn
}

/**
 * Account details
 *
 * @export
 */
export type IProviderAccount = {
  username?: string;
  id: string;
};
