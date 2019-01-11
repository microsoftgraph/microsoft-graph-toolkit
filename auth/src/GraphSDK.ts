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
}
