import { UserAgentApplication } from 'msal';
import { IAuthProvider, ProviderType, LoginChangedEvent, LoginType } from "./IAuthProvider";
import { Graph } from './GraphSDK';
import { EventHandler, EventDispatcher } from './EventHandler';
import { MSALConfig } from './MSALConfig';

export class MSALProvider implements IAuthProvider {

    private _loginChangedDispatcher = new EventDispatcher<LoginChangedEvent>();
    private _loginType : LoginType;
    private _idToken : string;

    readonly type: ProviderType;
    readonly provider: any;
    get isLogedIn() : boolean {
        return !!this._idToken;
    }

    scopes: string[];
    authority: string;
    
    graph: Graph;

    constructor(config: MSALConfig) {
        if (!config.clientId) {
            throw "ClientID must be a valid string";
        }

        this.scopes = (typeof config.scopes !== 'undefined') ? config.scopes : ["user.read"];
        this.authority = (typeof config.authority !== 'undefined') ? config.authority : null;
        let options = (typeof config.options != 'undefined') ? config.options : {};
        this._loginType = (typeof config.loginType !== 'undefined') ? config.loginType : LoginType.Redirect;

        this.type = ProviderType.V2;
        this.provider = new UserAgentApplication(config.clientId, this.authority, this.tokenReceivedCallback, options);
        this.graph = new Graph(this);
    }
    
    async login(): Promise<void> {
        let provider = this.provider as UserAgentApplication;
        
        if (this._loginType == LoginType.Redirect) {
            provider.loginRedirect(this.scopes);
        } else {
            this._idToken = await provider.loginPopup(this.scopes);
            this.fireLoginChangedEvent({});
        }
    }

    async getAccessToken(scopes?: string[]): Promise<string> {
        let provider = this.provider as UserAgentApplication;
        let accessToken : string;
        try {
            accessToken = await provider.acquireTokenSilent(scopes);
        } catch (e) {
            try {
                // TODO - figure out for what error this logic is needed so we
                // don't prompt the user to login unnecessarily
                if (this._loginType == LoginType.Redirect) {
                    await provider.acquireTokenRedirect(scopes);
                } else {
                    accessToken = await provider.acquireTokenPopup(scopes);
                }
            } catch (e) {
                // TODO - figure out how to expose this during dev to make it easy for the dev to figure out
                // if error contains "'token' is not enabled", make sure to have implicit oAuth enabled in the AAD manifest
                console.log(e);
                throw e;
            }
        }
        return accessToken;
    }
    
    async logout(): Promise<void> {
        let provider = this.provider as UserAgentApplication;
        provider.logout();
        this.fireLoginChangedEvent({});
    }
    
    updateScopes(scopes: string[]) {
        this.scopes = scopes;
    }

    tokenReceivedCallback(errorDesc : string, token: string, error: any, state: any)
    {
        if (error) {
            console.log(errorDesc);
        } else {
            this._idToken = token;
            this.fireLoginChangedEvent({});
        }
    }

    onLoginChanged(eventHandler : EventHandler<LoginChangedEvent>) {
        this._loginChangedDispatcher.register(eventHandler);
    }

    private fireLoginChangedEvent(event : LoginChangedEvent) {
        this._loginChangedDispatcher.fire(event);
    }
}