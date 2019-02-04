
import { MSALProvider } from './MSALProvider';
import { MSALConfig } from "./MSALConfig";
import { IAuthProvider, LoginType } from './IAuthProvider';
import { EventDispatcher, EventHandler } from './EventHandler';

let _provider : IAuthProvider = null;

export function getAuthProvider()
{
    return _provider;
}

export function initMSALProvider(clientId : string,
                                scopes? : string[],
                                authority? : string,
                                loginType? : LoginType)
{
    let config = {
        clientId: clientId,
        scopes: scopes,
        authority: authority,
        loginType : loginType
    }
    _provider = new MSALProvider(config);
    _eventDispatcher.fire( { newProvider: _provider } );
}

interface AuthProviderChangedEvent { newProvider : IAuthProvider }

let _eventDispatcher = new EventDispatcher<AuthProviderChangedEvent>();

export function onAuthProviderChanged(event : EventHandler<AuthProviderChangedEvent>) {
    _eventDispatcher.register(event)
}