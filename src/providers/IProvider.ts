import { IGraph } from "./Graph";
import { EventHandler } from "./EventHandler";

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

export interface LoginChangedEvent { }

export enum LoginType
{
    Popup,
    Redirect
}