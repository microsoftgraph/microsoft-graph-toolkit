import { MSALProvider as MsalProvider } from './MSALProvider';
import { MSALConfig } from "./MSALConfig";
import { IAuthProvider } from './IAuthProvider';
import { EventDispatcher, EventHandler } from './EventHandler';
import { WAMProvider } from './WAMProvider';

declare global {
    interface Window {
        _m365_providers : IAuthProvider[];
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

    export function add(provider : IAuthProvider) {
        const providers = getProviders();

        if (provider !== null) {
            providers.push(provider);
            _eventDispatcher.fire( {} );
        }
    }

    export function addWamProvider(clientId: string, authority?: string) {
        add(new WAMProvider(clientId, authority));
    }

    export function addMsalProvider(config : MSALConfig) {
        add(new MsalProvider(config));
    }

    let _eventDispatcher = new EventDispatcher();

    export function onProvidersChanged(event : EventHandler<any>) {
        _eventDispatcher.register(event)
    }

    // TODO - figure out a better way to have a global reference to all providers
    function getProviders() {
        if (!window._m365_providers) {
            window._m365_providers = [];
        }

        return window._m365_providers;
    }
}

export * from "./MSALConfig"
export * from "./MSALProvider"
export * from "./WAMProvider"
export * from "./IAuthProvider"
export * from "./GraphSDK"
export * from "./EventHandler"
