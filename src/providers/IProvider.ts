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
        console.log("SetState: " + state.toString());
    }

    // events
    onStateChanged(eventHandler: EventHandler<LoginChangedEvent>) {
        this._loginChangedDispatcher.register(eventHandler);
    }
    
    constructor() {
        this._state = ProviderState.Loading;
    }

    login?() : Promise<void>;
    logout?() : Promise<void>;
    abstract getAccessToken(...scopes: string[]) : Promise<string>;

    // get access to underlying provider
    provider : any;
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
    register(eventHandler: EventHandler<E>) {
        this.eventHandlers.push(eventHandler);
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