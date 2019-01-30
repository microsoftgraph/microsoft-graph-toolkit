import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"

export class Graph {

    private token: string;

    private rootUrl: string = 'https://graph.microsoft.com/beta';

    constructor(token: string) {
        this.token = token;
    }

    async getJson(resource: string) {
        let response = await this.get(resource);
        return response.json();
    }

    async get(resource: string) : Promise<Response> {
        return fetch(this.rootUrl + resource, {
            headers: {
                authorization: 'Bearer ' + this.token
            }
        });
    }

    async me() : Promise<MicrosoftGraph.User>
    {
        return this.getJson('/me') as MicrosoftGraph.User;
    }

    async photo() : Promise<Response> {
        return this.get('/me/photo/$value');
    }

    async calendar(startDateTime : Date, endDateTime : Date) : Promise<Array<MicrosoftGraph.Event>> {
        let sdt = `startdatetime=${startDateTime.toISOString()}`;
        let edt = `enddatetime=${endDateTime.toISOString()}`
        let uri = `/me/calendarview?${sdt}&${edt}`;
        let calendar = await this.getJson(uri);
        console.log(calendar);
        return calendar.value;
    }
}
