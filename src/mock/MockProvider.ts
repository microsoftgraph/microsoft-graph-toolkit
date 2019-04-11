import { IProvider, EventDispatcher, LoginChangedEvent, EventHandler, ProviderState } from "../providers/IProvider";
import { IGraph, Graph } from "../Graph";

export class MockProvider extends IProvider {

    constructor(signedIn: boolean = false) {
        super();
        if (signedIn) {
            this.setState(ProviderState.SignedIn);
        } else {
            this.setState(ProviderState.SignedOut);
        }
    }
    
    async login(): Promise<void> {
        this.setState(ProviderState.Loading);
        await (new Promise(resolve => setTimeout(resolve, 3000)));
        this.setState(ProviderState.SignedIn);
    }
    
    async logout(): Promise<void> {
        this.setState(ProviderState.Loading);
        await (new Promise(resolve => setTimeout(resolve, 3000)));
        this.setState(ProviderState.SignedOut);
    }
    
    getAccessToken(): Promise<string> {
        return Promise.resolve("");
    }
    provider: any;

    graph: IGraph = new MockGraph();
}

export class MockGraph extends Graph {

    private static baseUrl = "https://proxy.apisandbox.msdn.microsoft.com/svc?url=";
    private static rootGraphUrl: string = 'https://graph.microsoft.com/beta';

    constructor() {
        super(null);
    }

    async get(resource: string, scopes?: string[]) : Promise<Response> {
        if (!resource.startsWith('/')){
            resource = "/" + resource;
        }

        let response = await fetch(MockGraph.baseUrl + escape(MockGraph.rootGraphUrl + resource), {
            headers: {
                authorization: 'Bearer {token:https://graph.microsoft.com/}'
            }
        });

        if (response.status >= 400) {
            throw 'error accessing mock graph data';
        }

        return response;
    }

    async patch(resource: string, scopes: string[], data: any = null ): Promise<Response>
    {
        if (!resource.startsWith('/')){
            resource = "/" + resource;
        }

        let body = !!data ? {body: data}: {};
        let req = {
            url: MockGraph.baseUrl + escape(MockGraph.rootGraphUrl + resource),
            method: "PATCH",
            ...body,
        } as RequestInfo;

        console.log("Patch:", req);

        let response = await fetch(req, {
            headers: {
                authorization: 'Bearer {token:https://graph.microsoft.com/}'
            }
        });

        if (response.status >= 400) {
            throw 'error accessing mock graph data';
        }

        return response;
    }

    public async post(resource: string, scopes: string[], data: any)
    {
        if (!resource.startsWith('/')){
            resource = "/" + resource;
        }

        let body = !!data ? {body: data}: {};
        let req = {
            url: MockGraph.baseUrl + escape(MockGraph.rootGraphUrl + resource),
            method: "POST",
            ...body,
        } as RequestInfo;

        console.log("Post:", req);

        let response = await fetch(req, {
            headers: {
                authorization: 'Bearer {token:https://graph.microsoft.com/}'
            }
        });

        if (response.status >= 400) {
            throw 'error accessing mock graph data';
        }

        return response;
    }

    public async delete(resource: string, scopes: string[], data: any)
    {
        if (!resource.startsWith('/')){
            resource = "/" + resource;
        }

        let body = !!data ? {body: data}: {};
        let req = {
            url: MockGraph.baseUrl + escape(MockGraph.rootGraphUrl + resource),
            method: "DELETE",
            ...body,
        } as RequestInfo;

        console.log("Delete:", req);

        let response = await fetch(req, {
            headers: {
                authorization: 'Bearer {token:https://graph.microsoft.com/}'
            }
        });

        if (response.status >= 400) {
            throw 'error accessing mock graph data';
        }

        return response;
    }
}