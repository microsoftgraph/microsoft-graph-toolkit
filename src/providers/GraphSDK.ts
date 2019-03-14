import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"
import { IAuthProvider } from "./IAuthProvider";

export interface IGraph {
    me() : Promise<MicrosoftGraph.User>;
    getUser(id: string) : Promise<MicrosoftGraph.User>;
    findPerson(query: string) : Promise<MicrosoftGraph.Person[]>;
    myPhoto() : Promise<string>;
    getUserPhoto(id: string) : Promise<string>;
    calendar(startDateTime : Date, endDateTime : Date) : Promise<Array<MicrosoftGraph.Event>>
}

export class Graph implements IGraph {

    // private token: string;
    private _provider : IAuthProvider;

    private rootUrl: string = 'https://graph.microsoft.com/beta';

    constructor(provider: IAuthProvider) {
        // this.token = token;
        this._provider = provider;
    }

    async getJson(resource: string, scopes? : string[]) {
        let response = await this.get(resource, scopes);
        if (response) {
            return response.json();
        }

        return null;
    }

    async get(resource: string, scopes?: string[]) : Promise<Response> {
        if (!resource.startsWith('/')){
            resource = "/" + resource;
        }
        
        let token : string;
        try {
            token = await this._provider.getAccessToken(...scopes);
        } catch (error) {
            console.log(error);
            return null;
        }
        
        if (!token) {
            return null;
        }

        let response = await fetch(this.rootUrl + resource, {
            headers: {
                authorization: 'Bearer ' + token
            }
        });

        if (response.status >= 400) {

            // hit limit - need to wait and retry per:
            // https://docs.microsoft.com/en-us/graph/throttling
            if (response.status == 429) {
                console.log('too many requests - wait ' + response.headers.get('Retry-After') + ' seconds');
                return null;
            }

            let error : any = response.json();
            if (error.error !== undefined) {
                console.log(error);
            }
            console.log(response);
            throw 'error accessing graph';
        }

        return response;
    }

    async me() : Promise<MicrosoftGraph.User> {
        let scopes = ['user.read'];
        return this.getJson('/me', scopes) as MicrosoftGraph.User;
    }

    async getUser(userPrincipleName: string) : Promise<MicrosoftGraph.User> {
        let scopes = ['user.readbasic.all'];
        return this.getJson(`/users/${userPrincipleName}`, scopes) as MicrosoftGraph.User;
    }

    async findPerson(query: string) : Promise<MicrosoftGraph.Person[]>{
        let scopes = ['people.read'];
        let result = await this.getJson(`/me/people/?$search="${query}"`, scopes);
        return result ? result.value as MicrosoftGraph.Person[] : null;
    }

    myPhoto() : Promise<string> {
        let scopes = ['user.read'];
        return this.getBase64('/me/photo/$value', scopes);
    }

    async getUserPhoto(id: string) : Promise<string> {
        let scopes = ['user.readbasic.all'];
        return this.getBase64(`users/${id}/photo/$value`, scopes);
    }

    private async getBase64(resource: string, scopes: string[]) : Promise<string> {
        try {
            let response = await this.get(resource, scopes);
            if (!response) {
                return null;
            }

            let blob = await response.blob();
            
            return new Promise((resolve, reject) => {
                const reader = new FileReader;
                reader.onerror = reject;
                reader.onload = _ => {
                    resolve(reader.result as string);
                }
                reader.readAsDataURL(blob);
            });
        } catch {
            return null;
        }
    }

    async calendar(startDateTime : Date, endDateTime : Date) : Promise<Array<MicrosoftGraph.Event>> {
        let scopes = ['calendars.read'];

        let sdt = `startdatetime=${startDateTime.toISOString()}`;
        let edt = `enddatetime=${endDateTime.toISOString()}`
        let uri = `/me/calendarview?${sdt}&${edt}`;
        let calendar = await this.getJson(uri, scopes);
        return calendar ? calendar.value : null;
    }
}
