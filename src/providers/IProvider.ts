/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { AuthenticationProvider } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProvider';
import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';
import { Graph } from '../Graph';

export abstract class IProvider implements AuthenticationProvider {
  private _state: ProviderState;
  private _loginChangedDispatcher = new EventDispatcher<LoginChangedEvent>();

  public get state(): ProviderState {
    return this._state;
  }

  public setState(state: ProviderState) {
    if (state !== this._state) {
      this._state = state;
      this._loginChangedDispatcher.fire({});
    }
  }

  // events
  public onStateChanged(eventHandler: EventHandler<LoginChangedEvent>) {
    this._loginChangedDispatcher.add(eventHandler);
  }

  public removeStateChangedHandler(eventHandler: EventHandler<LoginChangedEvent>) {
    this._loginChangedDispatcher.remove(eventHandler);
  }

  constructor() {
    this._state = ProviderState.Loading;
  }

  public login?(): Promise<void>;
  public logout?(): Promise<void>;
  public getAccessTokenForScopes(...scopes: string[]): Promise<string> {
    return this.getAccessToken({ scopes });
  }

  public abstract getAccessToken(options?: AuthenticationProviderOptions): Promise<string>;
  public graph: Graph;
}

export type EventHandler<E> = (event: E) => void;
export interface LoginChangedEvent {}

export class EventDispatcher<E> {
  private eventHandlers: Array<EventHandler<E>> = [];

  public fire(event: E) {
    for (const handler of this.eventHandlers) {
      handler(event);
    }
  }

  public add(eventHandler: EventHandler<E>) {
    this.eventHandlers.push(eventHandler);
  }

  public remove(eventHandler: EventHandler<E>) {
    for (let i = 0; i < this.eventHandlers.length; i++) {
      if (this.eventHandlers[i] === eventHandler) {
        this.eventHandlers.splice(i, 1);
        i--;
      }
    }
  }
}

export enum LoginType {
  Popup,
  Redirect
}

export enum ProviderState {
  Loading,
  SignedOut,
  SignedIn
}
