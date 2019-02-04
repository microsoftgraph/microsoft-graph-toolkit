import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"
import { IAuthProvider } from "./IAuthProvider";

export class Graph {

    // private token: string;
    private _provider : IAuthProvider;

    private rootUrl: string = 'https://graph.microsoft.com/beta';

    constructor(provider: IAuthProvider) {
        // this.token = token;
        this._provider = provider;
    }

    async getJson(resource: string, scopes? : string[]) {
        let response = await this.get(resource, scopes);
        return response.json();
    }

    async get(resource: string, scopes?: string[]) : Promise<Response> {
        let token : string;
        try {
            if (typeof scopes !== 'undefined') {
                token = await this._provider.getAccessToken(scopes);
            } else {
                token = await this._provider.getAccessToken();
            }
        } catch (error) {
            console.log(error);
            throw 'Unable to retreive token for Graph';
        }
        

        let response = await fetch(this.rootUrl + resource, {
            headers: {
                authorization: 'Bearer ' + token
            }
        });

        if (response.status >= 400) {
            let error : any = response.json();
            if (error.error !== undefined) {
                console.log(error);
            }
            throw 'error accessing graph';
        }

        return response;
    }

    async me() : Promise<MicrosoftGraph.User>
    {
        let scopes = ['user.read'];
        return this.getJson('/me', scopes) as MicrosoftGraph.User;
    }

    async photo() : Promise<Response> {
        let scopes = ['user.read'];
        return this.get('/me/photo/$value', scopes);
    }

    async calendar(startDateTime : Date, endDateTime : Date) : Promise<Array<MicrosoftGraph.Event>> {
        let scopes = ['calendars.read'];

        let sdt = `startdatetime=${startDateTime.toISOString()}`;
        let edt = `enddatetime=${endDateTime.toISOString()}`
        let uri = `/me/calendarview?${sdt}&${edt}`;
        let calendar = await this.getJson(uri, scopes);
        return calendar.value;
    }
}
