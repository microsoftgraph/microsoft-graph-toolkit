/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { AuthenticationProvider, AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client';
import { validateBaseURL } from '../utils/GraphHelpers';
import { GraphEndpoint, IGraph, MICROSOFT_GRAPH_DEFAULT_ENDPOINT } from '../IGraph';
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
   * Specifies if the provider has enabled support for multiple accounts
   *
   * @protected
   * @type {boolean}
   * @memberof IProvider
   */
  protected isMultipleAccountDisabled = true;

  /**
   * Specifies if Multi account functionality is supported by the provider and enabled.
   *
   * @readonly
   * @type {boolean}
   * @memberof IProvider
   */
  public get isMultiAccountSupportedAndEnabled(): boolean {
    return false;
  }
  private _state: ProviderState;
  private readonly _loginChangedDispatcher = new EventDispatcher<LoginChangedEvent>();
  private readonly _activeAccountChangedDispatcher = new EventDispatcher<ActiveAccountChanged>();
  private _baseURL: GraphEndpoint = MICROSOFT_GRAPH_DEFAULT_ENDPOINT;

  /**
   * The base URL to be used in the graph client config.
   */
  public set baseURL(url: GraphEndpoint) {
    if (validateBaseURL(url)) {
      this._baseURL = url;
      return;
    } else {
      throw new Error(`${url} is not a valid Graph URL endpoint.`);
    }
  }

  public get baseURL(): GraphEndpoint {
    return this._baseURL;
  }

  private _customHosts?: string[] = undefined;

  /**
   * Custom Hostnames to allow graph client to utilize
   */
  public set customHosts(hosts: string[] | undefined) {
    this._customHosts = hosts;
  }

  public get customHosts(): string[] | undefined {
    return this._customHosts;
  }

  /**
   * Enable/Disable incremental consent
   *
   * @protected
   * @type {boolean}
   * @memberof IProvider
   */
  private _isIncrementalConsentDisabled = false;

  /**
   * Backing field for isMultiAccountSupported
   *
   * @protected
   * @memberof IProvider
   */
  protected isMultipleAccountSupported = false;

  /**
   * Does the provider support multiple accounts?
   *
   * @readonly
   * @type {boolean}
   * @memberof IProvider
   */
  public get isMultiAccountSupported(): boolean {
    return this.isMultipleAccountSupported;
  }
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
   * Returns active account in case of multi-account sign in
   *
   * @return {*}  {any[]}
   * @memberof IProvider
   */
  public getActiveAccount?(): IProviderAccount;

  /**
   * Switch between two signed in accounts
   *
   * @param {*} user
   * @memberof IProvider
   */
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
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
// eslint-disable-next-line @typescript-eslint/no-empty-interface
export interface ActiveAccountChanged {}
/**
 * loginChangedEvent
 *
 * @export
 * @interface LoginChangedEvent
 */
// eslint-disable-next-line @typescript-eslint/no-empty-interface
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
export interface IProviderAccount {
  id: string;
  mail?: string;
  name?: string;
  tenantId?: string;
}
