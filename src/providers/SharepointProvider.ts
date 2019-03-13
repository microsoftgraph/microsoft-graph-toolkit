import { IAuthProvider, LoginChangedEvent, LoginType } from "./IAuthProvider";
import { IGraph, Graph } from './GraphSDK';
import { EventHandler, EventDispatcher } from './EventHandler';

declare interface AadTokenProvider{
    getToken(x:string);
}

export declare interface WebPartContext{
    aadTokenProviderFactory : any;
}

export class SharepointProvider implements IAuthProvider {
    
    private _loginChangedDispatcher = new EventDispatcher<LoginChangedEvent>();
    
    private _idToken : string;

    private _provider : AadTokenProvider;
    
    get provider() {
        return this._provider;
    };

    get isLoggedIn() : boolean {
        return !!this._idToken;
    };

    get isAvailable(): boolean{
        return true;
    };

    private context : WebPartContext;

    scopes: string[];
    authority: string;
    
    graph: IGraph;

    constructor(context : WebPartContext) {

        this.context = context;

        context.aadTokenProviderFactory.getTokenProvider().then((tokenProvider: AadTokenProvider): void => {
            this._provider = tokenProvider;
            this.graph = new Graph(this);
            this.login();
          });
        this.fireLoginChangedEvent({});
    }
    
    async login(): Promise<void> {
        this._idToken = await this.getAccessToken();
        if (this._idToken) {
            this.fireLoginChangedEvent({});
        }
    }

    async getAccessToken(): Promise<string> {
        let accessToken : string;
        try {
            accessToken = await this.provider.getToken("https://graph.microsoft.com");
        } catch (e) {
            console.log(e);
            throw e;
        }
        return accessToken;
    }
    
    async logout(): Promise<void> {
        //this.provider.logout();
        this.fireLoginChangedEvent({});
    }
    
    updateScopes(scopes: string[]) {
        this.scopes = scopes;
    }

    onLoginChanged(eventHandler : EventHandler<LoginChangedEvent>) {
        this._loginChangedDispatcher.register(eventHandler);
    }

    private fireLoginChangedEvent(event : LoginChangedEvent) {
        this._loginChangedDispatcher.fire(event);
    }
}