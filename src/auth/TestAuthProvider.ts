import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"
import { IAuthProvider, LoginChangedEvent } from "./IAuthProvider";
import { IGraph } from "./GraphSDK";
import { EventDispatcher } from "./EventHandler";

export class TestAuthProvider implements IAuthProvider {

    isLoggedIn: boolean = false;
    private _loginChangedDispatcher = new EventDispatcher<LoginChangedEvent>();

    login(): Promise<void> {
        this.isLoggedIn = true;
        this._loginChangedDispatcher.fire({});
        return Promise.resolve();
    }
    logout(): Promise<void> {
        this.isLoggedIn = false;
        this._loginChangedDispatcher.fire({});
        return Promise.resolve();
    }
    getAccessToken(scopes?: string[]): Promise<string> {
        return Promise.resolve("");
    }
    provider: any;

    graph: IGraph = new TestGraph();

    onLoginChanged(eventHandler: import("c:/Users/nikol/source/M365Toolkit/src/auth/EventHandler").EventHandler<import("c:/Users/nikol/source/M365Toolkit/src/auth/IAuthProvider").LoginChangedEvent>) {
        this._loginChangedDispatcher.register(eventHandler);
    }
}

export class TestGraph implements IGraph {
    calendar(startDateTime: Date, endDateTime: Date): Promise<MicrosoftGraph.Event[]> {
        let calendarData : MicrosoftGraph.Event[] = [
            {
                subject : 'event 1'
            },
            {
                subject : 'event 2'
            }
        ]

        return Promise.resolve(calendarData);
    }

    me() : Promise<MicrosoftGraph.User> {
        let testUser : MicrosoftGraph.User =  {
            displayName: 'Test User',
            mail: 'email@contoso.com'
        }
        return Promise.resolve(testUser);
    }

    myPhoto(): Promise<Response> {
        return fetch('https://pbs.twimg.com/profile_images/861791187019079684/_-blnGB8_400x400.jpg');
    }

}