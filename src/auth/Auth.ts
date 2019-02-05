
import { MSALProvider } from './MSALProvider';
import { MSALConfig } from "./MSALConfig";
import { IAuthProvider, LoginType } from './IAuthProvider';
import { EventDispatcher, EventHandler } from './EventHandler';
import { TestAuthProvider } from './TestAuthProvider';

let _provider : IAuthProvider = null;

let _useFakeAuth : boolean = true;

export function getAuthProvider()
{
    return _provider;
}

export function useFakeAuth(){
    _useFakeAuth = true;
}

export function init(clientId : string, loginType? : LoginType) {
    if(_useFakeAuth) {
        _provider = new TestAuthProvider();
    } else {
        let config : MSALConfig = {
            clientId: clientId,
            loginType: loginType
        };
        _provider = new MSALProvider(config);
    }

    _eventDispatcher.fire( { newProvider: _provider } );
}

interface AuthProviderChangedEvent { newProvider : IAuthProvider }

let _eventDispatcher = new EventDispatcher<AuthProviderChangedEvent>();

export function onAuthProviderChanged(event : EventHandler<AuthProviderChangedEvent>) {
    _eventDispatcher.register(event)
}