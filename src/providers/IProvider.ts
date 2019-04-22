import { IGraph } from '../Graph';
import { AuthenticationProvider } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProvider';
import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';

export abstract class IProvider implements AuthenticationProvider {
  private _state: ProviderState;
  private _loginChangedDispatcher = new EventDispatcher<LoginChangedEvent>();

  get state(): ProviderState {
    return this._state;
  }

  setState(state: ProviderState) {
    if (state != this._state) {
      this._state = state;
      this._loginChangedDispatcher.fire({});
    }
  }

  // events
  onStateChanged(eventHandler: EventHandler<LoginChangedEvent>) {
    this._loginChangedDispatcher.add(eventHandler);
  }

  removeStateChangedHandler(eventHandler: EventHandler<LoginChangedEvent>) {
    this._loginChangedDispatcher.remove(eventHandler);
  }

  constructor() {
    this._state = ProviderState.Loading;
  }

  login?(): Promise<void>;
  logout?(): Promise<void>;
  getAccessTokenForScopes(...scopes: string[]): Promise<string> {
    return this.getAccessToken({ scopes: scopes });
  }

  abstract getAccessToken(options?: AuthenticationProviderOptions): Promise<string>;
  graph: IGraph;
}

export type EventHandler<E> = (event: E) => void;
export interface LoginChangedEvent {}

export class EventDispatcher<E> {
  private eventHandlers: EventHandler<E>[] = [];

  fire(event: E) {
    for (let handler of this.eventHandlers) {
      handler(event);
    }
  }

  add(eventHandler: EventHandler<E>) {
    this.eventHandlers.push(eventHandler);
  }

  remove(eventHandler: EventHandler<E>) {
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
