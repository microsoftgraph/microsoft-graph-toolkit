import { UserAgentApplication } from 'msal';
import { IAuthProvider, ProviderType } from "./IAuthProvider";
import { Graph } from './GraphSDK';

export class V2Provider implements IAuthProvider {
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
            let accessToken = await provider.acquireTokenSilent(this.scopes);
            if (accessToken) {
                this.graph = new Graph(accessToken);
            }
        }
    }
    
    async logout(): Promise<void> {
        let provider = this.provider as UserAgentApplication;
        provider.logout();
    }
    
    updateScopes(scopes: string[]) {
        this.scopes = scopes;
    }

    tokenReceivedCallback(token : any)
    {
        return token;
    }
}
