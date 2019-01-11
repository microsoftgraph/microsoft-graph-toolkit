
import { V2Provider } from './V2Provider';

let _provider = null;

export function getAuthProvider()
{
    return _provider;
}

export function initV2Provider(clientId : string,
                                scopes : string[] = ["user.read"],
                                authority : string = null,
                                options : any = null)
{
    _provider = new V2Provider(clientId, scopes, authority, options);
}