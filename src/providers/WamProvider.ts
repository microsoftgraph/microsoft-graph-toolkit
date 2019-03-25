import { IProvider, LoginChangedEvent } from "./IProvider";
import { IGraph, Graph } from "./Graph";
import { EventHandler, EventDispatcher } from "./EventHandler";

declare interface Window {
    Windows: any;
}

declare var window: Window;

export class WamProvider implements IProvider {

    private graphResource = 'https://graph.microsoft.com';
    private clientId : string;
    private authority : string;

    private accessToken : string;

    private _loginChangedDispatcher = new EventDispatcher<LoginChangedEvent>();

    public static isAvailable() : boolean {
        return !!(window.Windows);
    }

    get isLoggedIn() : boolean{
        return !!this.accessToken;
    }

    constructor(clientId: string, authority?: string) {
        this.clientId = clientId;
        this.authority = authority || 'https://login.microsoftonline.com/common';

        this.graph = new Graph(this);

        this.printRedirectUriToConsole();
    }

    
    async login(): Promise<void> {
        if (WamProvider.isAvailable()) {

            let webCore = window.Windows.Security.Authentication.Web.Core;
            let wap = await webCore.WebAuthenticationCoreManager.findAccountProviderAsync('https://login.microsoft.com', this.authority);
            if (!wap) {
                console.log('no account provider');
                return;
            }

            let wtr = new webCore.WebTokenRequest(wap, '', this.clientId);
            wtr.properties.insert('resource', this.graphResource);
            
            let wtrr = await webCore.WebAuthenticationCoreManager.requestTokenAsync(wtr);
            switch (wtrr.responseStatus) {
                case webCore.WebTokenRequestStatus.success:
                    let account = wtrr.responseData[0].webAccount;
                    this.accessToken = wtrr.responseData[0].token;
                    this.fireLoginChangedEvent({});
                    break;
                case webCore.WebTokenRequestStatus.userCancel:
                case webCore.WebTokenRequestStatus.accountSwitch:
                case webCore.WebTokenRequestStatus.userInteractionRequired:
                case webCore.WebTokenRequestStatus.accountProviderNotAvailable:
                case webCore.WebTokenRequestStatus.providerError:
                    console.log(`status ${wtrr.responseStatus}: error code ${wtrr.responseError} | error message ${wtrr.responseError.errorMessage}`);
                    break;
            }
        }
    }

    printRedirectUriToConsole() {
        if (WamProvider.isAvailable()) {
            let web = window.Windows.Security.Authentication.Web;
            let redirectUri = `ms-appx-web://Microsoft.AAD.BrokerPlugIn/${(web.WebAuthenticationBroker.getCurrentApplicationCallbackUri().host as string).toUpperCase()}`;
            console.log("Use the following redirect URI in your AAD application:")
            console.log(redirectUri);
        } else {
            console.log("WAM not supported on this platform");
        }
    }

    logout(): Promise<void> {
        throw new Error("Method not implemented.");
    }

    getAccessToken(...scopes: string[]): Promise<string> {
        if (this.isLoggedIn) {
            return Promise.resolve(this.accessToken);
        }
        else {
            throw "Not logged in";
        }

    }

    provider: any;

    graph: IGraph;

    onLoginChanged(eventHandler : EventHandler<LoginChangedEvent>) {
        this._loginChangedDispatcher.register(eventHandler);
    }

    private fireLoginChangedEvent(event : LoginChangedEvent) {
        this._loginChangedDispatcher.fire(event);
    }
}