
import { MSALProvider } from './MSALProvider';
import { MSALConfig } from "./MSALConfig";
import { IAuthProvider, LoginType } from './IAuthProvider';
import { EventDispatcher, EventHandler } from './EventHandler';
import { WAMProvider } from './WAMProvider';
import { TestAuthProvider } from './TestAuthProvider';

let _provider : IAuthProvider = null;

export function getAuthProvider()
{
    return _provider;
}

export function initWithProvider(provider : IAuthProvider) {
    _provider = provider;
    _eventDispatcher.fire( { newProvider: _provider } );
}

export function initWithWam(clientId: string, authority?: string) {
    _provider = new WAMProvider(clientId, authority);
    _eventDispatcher.fire( { newProvider: _provider } );
}

export function initWithCustomProvider(provider: IAuthProvider) {
    _provider = provider;
    _eventDispatcher.fire( { newProvider: _provider } );
}

export function initWithFakeProvider() {
    initWithProvider(new TestAuthProvider());
}

export function initMSALProvider(config : MSALConfig) {
    initWithProvider(new MSALProvider(config));
}

interface AuthProviderChangedEvent { newProvider : IAuthProvider }

let _eventDispatcher = new EventDispatcher<AuthProviderChangedEvent>();

export function onAuthProviderChanged(event : EventHandler<AuthProviderChangedEvent>) {
    _eventDispatcher.register(event)
}