
import { V2Provider } from './V2Provider';
import { IAuthProvider } from './IAuthProvider';
import { EventDispatcher, EventHandler } from './EventHandler';

let _provider : IAuthProvider = null;

export function getAuthProvider()
{
    return _provider;
}

export function initV2Provider(clientId : string,
                                scopes : string[] = ["user.read", "calendars.read"],
                                authority : string = null,
                                options : any = { cacheLocation: 'localStorage'})
{
    _provider = new V2Provider(clientId, scopes, authority, options);
    _eventDispatcher.fire( { newProvider: _provider } );
}

interface AuthProviderChangedEvent { newProvider : IAuthProvider }

let _eventDispatcher = new EventDispatcher<AuthProviderChangedEvent>();

export function onAuthProviderChanged(event : EventHandler<AuthProviderChangedEvent>) {
    _eventDispatcher.register(event)
}