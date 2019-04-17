import { IGraph } from "../Graph";

export abstract class IProvider
{
    private _state : ProviderState;
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

    login?() : Promise<void>;
    logout?() : Promise<void>;
    abstract getAccessToken(...scopes: string[]) : Promise<string>;

    graph : IGraph;
}

export type EventHandler<E> = (event: E) => void;
export interface LoginChangedEvent { }

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
            if (this.eventHandlers[i] === eventHandler){
                this.eventHandlers.splice(i, 1);
                i--;
            }
        }
    }
}

export enum LoginType
{
    Popup,
    Redirect
}

export enum ProviderState
{
    Loading,
    SignedOut,
    SignedIn
}