import { IGraph } from "../Graph";

export interface IProvider 
{
    readonly isLoggedIn : boolean;
    
    login?() : Promise<void>;
    logout?() : Promise<void>;
    getAccessToken(...scopes: string[]) : Promise<string>;

    // get access to underlying provider
    provider : any;
    graph : IGraph;

    // events
    onLoginChanged(eventHandler : EventHandler<LoginChangedEvent>)
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