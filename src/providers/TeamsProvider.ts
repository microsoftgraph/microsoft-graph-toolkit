import { IAuthProvider, LoginChangedEvent } from "./IAuthProvider";
import { IGraph, Graph } from './GraphSDK';
import { EventHandler, EventDispatcher } from './EventHandler';
import * as microsoftTeams from "@microsoft/teams-js";

export class TeamsProvider implements IAuthProvider {
    
    private _loginChangedDispatcher = new EventDispatcher<LoginChangedEvent>();
    
    private _idToken : string;

    private _clientId : string;
    private _loginPopupUrl? : string;
    private _loginPopupEndUrl? : string;

    private _provider : any;
    
    get provider() {
        return this._provider;
    };

    get isLoggedIn() : boolean {
        return !!this._idToken;
    };

    get clientId() : string {
        return this._clientId;
    }

    set clientId(value: string) {
        this._clientId = value;
    }

    get loginPopupUrl() : string {
        return this._loginPopupUrl;
    }

    set loginPopupUrl(value: string) {
        this._loginPopupUrl = value;
    }

    get loginPopupEndUrl() : string {
        return this._loginPopupEndUrl;
    }

    set loginPopupEndUrl(value: string) {
        this._loginPopupEndUrl = value;
    }

    static async isAvailable(): Promise<boolean>{
        const ms = 500;
        return Promise.race([
            new Promise<boolean>((resolve, reject)=>{
                try{
                    microsoftTeams.initialize();
                    microsoftTeams.getContext(function(context){
                        if(context){
                            resolve(true);
                        }else{
                            resolve(false);
                        }
                    });
                }catch(reason){
                    resolve(false);
                }
            }),
            new Promise<boolean>((resolve, reject) => {
                let id = setTimeout(() => {
                  clearTimeout(id);
                  resolve(false);
                }, ms)
            })
        ]);
    };

    scopes: string[];
    authority: string;
    
    graph: IGraph;

    constructor(clientId: string, loginPopupUrl?: string, loginPopupEndUrl?: string) {

        this._clientId = clientId;
        this._loginPopupUrl = loginPopupUrl;
        this._loginPopupEndUrl = loginPopupEndUrl;
        
        this._provider = this;
        this.graph = new Graph(this);
    }
    
    async login(): Promise<void> {
        this._idToken = await this.getAccessToken();
        if (this._idToken) {
            this.fireLoginChangedEvent({});
        }
    }

    async getAccessToken(): Promise<string> {
        if (this._idToken)
            return this._idToken;
        return new Promise((resolve, reject) => {
            var loginPopupUrl = this._loginPopupUrl;
            if(!loginPopupUrl){
                loginPopupUrl = window.location.origin + "/auth.html";
            }

            var loginPopupEndUrl = this._loginPopupEndUrl;
            if(!loginPopupEndUrl){
                loginPopupEndUrl = window.location.origin + "/auth-end.html";
            }
            
            var url = new URL(loginPopupUrl);
            url.searchParams.append('loginPopupEndUrl', loginPopupEndUrl);
            url.searchParams.append('clientId', this._clientId);

            microsoftTeams.authentication.authenticate({
                url: url.href,
                width: 600,
                height: 535,
                successCallback: function (result: any) {
                    resolve(result.accessToken);
                },
                failureCallback: function (reason) {
                    console.log(reason);
                    reject(reason);
                }
            });
        });
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