import { IGraph } from "./GraphSDK";
import { EventHandler } from "./EventHandler";

export interface IAuthProvider 
{
    readonly isLoggedIn : boolean;
    readonly isAvailable : boolean;

    login() : Promise<void>;
    logout() : Promise<void>;
    getAccessToken() : Promise<string>;
    addScope?(...scope: string[]);

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