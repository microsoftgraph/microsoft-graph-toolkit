import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"

export class Graph {

    private token: string;

    private rootUrl: string = 'https://graph.microsoft.com/v1.0';

    constructor(token: string) {
        this.token = token;
    }

    async get(resource: string) {
        let response = await fetch(this.rootUrl + resource, {
            headers: {
                authorization: 'Bearer ' + this.token
            }
        });
        return response.json();
    }

    async me() : Promise<MicrosoftGraph.User>
    {
        return this.get('/me') as MicrosoftGraph.User;
    }

    async calendar(startDateTime : Date, endDateTime : Date) : Promise<Array<MicrosoftGraph.Event>> {
        let sdt = `startdatetime=${startDateTime.toISOString()}`;
        let edt = `enddatetime=${endDateTime.toISOString()}`
        let uri = `/me/calendarview?${sdt}&${edt}`;
        let calendar = await this.get(uri);
        console.log(calendar);
        return calendar.value;
    }
}
