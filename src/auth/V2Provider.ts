import { UserAgentApplication } from 'msal';
import { IAuthProvider, ProviderType, LoginChangedEvent } from "./IAuthProvider";
import { Graph } from './GraphSDK';
import { EventHandler, EventDispatcher } from './EventHandler';

export class V2Provider implements IAuthProvider {

    private _loginChangedDispatcher = new EventDispatcher<LoginChangedEvent>();

    readonly type: ProviderType;
    readonly provider: any;

    scopes: string[];
    
    graph: Graph;
    
    constructor(clientId: string, scopes: string[] = ["user.read"], authority: string = null, options: any = null) {
        this.type = ProviderType.V2;
        this.provider = new UserAgentApplication(clientId, authority, this.tokenReceivedCallback);
        this.scopes = scopes;
    }
    
    async login(): Promise<void> {
        let provider = this.provider as UserAgentApplication;
        let token = await provider.loginPopup(this.scopes);
        if (token) {
            await this.loginSilent();
        }
    }

    async loginSilent(): Promise<boolean> {
        let provider = this.provider as UserAgentApplication;
        try {
            let accessToken = await provider.acquireTokenSilent(this.scopes);
            if (accessToken) {
                this.graph = new Graph(accessToken);
                this.fireLoginChangedEvent({});
                return true;
            }
        } catch {
            return false;
        }

        return false;
    }
    
    async logout(): Promise<void> {
        let provider = this.provider as UserAgentApplication;
        provider.logout();
        this.fireLoginChangedEvent({});
    }
    
    updateScopes(scopes: string[]) {
        this.scopes = scopes;
    }

    tokenReceivedCallback(token : any)
    {
        return token;
    }

    onLoginChanged(eventHandler : EventHandler<LoginChangedEvent>) {
        this._loginChangedDispatcher.register(eventHandler);
    }

    private fireLoginChangedEvent(event : LoginChangedEvent) {
        this._loginChangedDispatcher.fire(event);
    }
}
