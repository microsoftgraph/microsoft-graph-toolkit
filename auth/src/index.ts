
import { V2Provider } from './V2Provider';
import { IAuthProvider } from './IAuthProvider';

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
}