import { MsalProvider as MsalProvider } from './MsalProvider';
import { MsalConfig } from "./MsalConfig";
import { IProvider } from './IProvider';
import { EventDispatcher, EventHandler } from './EventHandler';
import { WamProvider } from './WamProvider';
import { SharePointProvider, WebPartContext } from './SharePointProvider';

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
            if (provider.isAvailable)
                return provider;
        }
        return null;
    }

    export function add(provider : IProvider) {
        const providers = getProviders();

        if (provider !== null) {
            providers.push(provider);
            getEventDispatcher().fire( {} );
        }
    }

    export function addWamProvider(clientId: string, authority?: string) {
        add(new WamProvider(clientId, authority));
    }

    export function addMsalProvider(config : MsalConfig) {
        add(new MsalProvider(config));
    }

    export function addSharePointProvider(context : WebPartContext ) {
        add(new SharePointProvider(context));
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
export * from "./IProvider"
export * from "./Graph"
export * from "./EventHandler"
