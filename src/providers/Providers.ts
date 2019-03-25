import { MsalProvider as MsalProvider } from './MsalProvider';
import { MsalConfig } from "./MsalConfig";
import { IProvider } from './IProvider';
import { EventDispatcher, EventHandler } from './EventHandler';
import { WamProvider } from './WamProvider';
import { SharePointProvider, WebPartContext } from './SharePointProvider';
import { TeamsProvider } from './TeamsProvider';

declare global {
    interface Window {
        _msgraph_providers : IProvider[],
        _msgraph_eventDispatcher : EventDispatcher<{}>
    }
}

export module Providers {
    export function getAvailable() {
        const providers = getProviders();

        for (let provider of providers) {
            return provider;
        }
        return null;
    }

    export function addCustomProvider(provider : IProvider): IProvider {
        const providers = getProviders();

        if (provider !== null) {
            providers.push(provider);
            getEventDispatcher().fire( {} );
        }
        return provider;
    }

    export function addWamProvider(clientId: string, authority?: string) {
        if(WamProvider.isAvailable()){
            return <WamProvider>addCustomProvider(new WamProvider(clientId, authority));
        }
        return null;
    }

    export function addMsalProvider(config : MsalConfig) {
        return <MsalProvider>addCustomProvider(new MsalProvider(config));
    }

    export function addSharePointProvider(context : WebPartContext ) {
        return <SharePointProvider>addCustomProvider(new SharePointProvider(context));
    }

    export async function addTeamsProvider(clientId: string, loginPopupUrl: string) {
        if(await TeamsProvider.isAvailable()){
            return <TeamsProvider>addCustomProvider(new TeamsProvider(clientId, loginPopupUrl));
        }
        return null;
    }

    export function onProvidersChanged(event : EventHandler<any>) {
        getEventDispatcher().register(event)
    }

    // TODO - figure out a better way to have a global reference to all providers
    function getProviders() {
        if (!window._msgraph_providers) {
            window._msgraph_providers = [];
        }

        return window._msgraph_providers;
    }

    function getEventDispatcher() {
        if (!window._msgraph_eventDispatcher) {
            window._msgraph_eventDispatcher = new EventDispatcher();
        }

        return window._msgraph_eventDispatcher;
    }
}

export * from "./MsalConfig"
export * from "./MsalProvider"
export * from "./WamProvider"
export * from "./SharePointProvider"
export * from "./TeamsProvider"
export * from "./IProvider"
export * from "./Graph"
export * from "./EventHandler"
