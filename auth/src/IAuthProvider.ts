import { Graph } from "./GraphSDK";
import { EventHandler } from "./EventHandler";

export interface IAuthProvider 
{
    type : ProviderType;

    login() : Promise<void>;
    loginSilent() : Promise<void>;
    logout() : Promise<void>;

    // get access to underlying provider
    provider : any;
    graph : Graph;

    // events
    onLoginChanged(eventHandler : EventHandler<LoginChangedEvent>)
}

export interface LoginChangedEvent { }

export enum ProviderType
{
    V2
}

export enum LoginType
{
    Popup,
    Redirect
}