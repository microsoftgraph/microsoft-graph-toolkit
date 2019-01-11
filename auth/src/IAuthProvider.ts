import { Graph } from "./GraphSDK";

export interface IAuthProvider 
{
    type : ProviderType;

    login() : Promise<void>;
    logout() : Promise<void>;

    // get access to underlying provider
    provider : any;
    graph : Graph;
}

export enum ProviderType
{
    V2
}

export enum LoginType
{
    Popup,
    Redirect
}