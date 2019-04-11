import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"
import { IProvider } from "./providers/IProvider";

export interface IGraph {
    me() : Promise<MicrosoftGraph.User>;
    getUser(id: string) : Promise<MicrosoftGraph.User>;
    findPerson(query: string) : Promise<MicrosoftGraph.Person[]>;
    myPhoto() : Promise<string>;
    getUserPhoto(id: string) : Promise<string>;
    calendar(startDateTime : Date, endDateTime : Date) : Promise<Array<MicrosoftGraph.Event>>;

    // Planner Task Methods
    getAllMyPlans?(): Promise<MicrosoftGraph.PlannerPlan[]>;
    getSinglePlan?(planId: string): Promise<MicrosoftGraph.PlannerPlan>;
    getPlanDetails?(planId: string): Promise<MicrosoftGraph.PlannerPlanDetails>;

    getAllMyTasks?(): Promise<MicrosoftGraph.PlannerTask[]>;
    getTasksForPlan?(taskId: string): Promise<MicrosoftGraph.PlannerTask[]>;
    getTaskDetails?(taskId: string): Promise<MicrosoftGraph.PlannerTaskDetails>;

    setTaskDetails?(taskId: string, newInfo: MicrosoftGraph.PlannerTask): Promise<any>;
    setTaskComplete?(taskId: string): Promise<any>;
    addTask?(planId: string, newTask: MicrosoftGraph.PlannerTask): Promise<any>;
    removeTask?(taskId: string): Promise<any>;
}

export class Graph implements IGraph {

    // private token: string;
    private _provider : IProvider;

    private rootUrl: string = 'https://graph.microsoft.com/beta';

    constructor(provider: IProvider) {
        // this.token = token;
        this._provider = provider;
    }

    private checkResource(resource: string): string 
    {
        if (!resource.startsWith('/')){
            resource = "/" + resource;
        }

        return resource;
    }

    private checkResponse(response: Response)
    {
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
    }

    async getJson(resource: string, scopes? : string[]) {
        let response = await this.get(resource, scopes);
        if (response) {
            return response.json();
        }

        return null;
    }

    async get(resource: string, scopes?: string[]) : Promise<Response> {
        resource = this.checkResource(resource);

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
        this.checkResponse(response);
        return response;
    }

    async patch(resource: string, scopes: string[], data?: any): Promise<Response>
    {
        resource = this.checkResource(resource);
        
        let token = await this._provider.getAccessToken(...scopes).catch(err=>null);
        if(!token)
            return null;

        let body = !!data ? {body: data} : {};

        let result = await fetch({
            method: 'PATCH',
            url: this.rootUrl + resource,
            ...body
        } as RequestInfo,
        {
            headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": 'application/json'
        }});
        this.checkResponse(result);
        return result;
    }

    async post(resource: string, scopes: string[], data: any)
    {
        resource = this.checkResource(resource);
        
        let token = await this._provider.getAccessToken(...scopes).catch(err=>null);
        if(!token)
            return null;

        let body = !!data ? {body: data} : {};

        let result = await fetch({
            method: 'POST',
            url: this.rootUrl + resource,
            ...body
        } as RequestInfo,
        {
            headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": 'application/json'
        }});

        this.checkResponse(result);
        return result;
    }

    async delete(resource: string, scopes: string[], data: any = null)
    {
        resource = this.checkResource(resource);
        
        let token = await this._provider.getAccessToken(...scopes).catch(err=>null);
        if(!token)
            return null;

        let body = !!data ? {body: data} : {};

        let result = await fetch({
            method: 'DELETE',
            url: this.rootUrl + resource,
            ...body
        } as RequestInfo,
        {
            headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": 'application/json'
        }});

        this.checkResponse(result);
        return result;
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


    public async getAllMyPlans(): Promise<MicrosoftGraph.PlannerPlan[]>
    {
        let scopes = ['Group.ReadWrite.All'];
        let planners = await this.getJson('/me/planner/plans', scopes) as {value: MicrosoftGraph.PlannerPlan[]};
        return planners.value;
    }

    public async getSinglePlan(planId: string): Promise<MicrosoftGraph.PlannerPlan>
    {
        let scopes = ['Group.ReadWrite.All'];
        let plan = await this.getJson(`/planner/plans/${planId}`, scopes) as MicrosoftGraph.PlannerPlan;

        return plan;
    }

    public async getTasksForPlan(planId: string): Promise<MicrosoftGraph.PlannerTask[]>
    {
        let scopes = ['Group.ReadWrite.All'];
        let tasks = await this.getJson(`/planner/plans/${planId}/tasks`, scopes) as {value: MicrosoftGraph.PlannerTask[]};
        return tasks.value;
    }

    public async getTaskDetails(taskId: string): Promise<MicrosoftGraph.PlannerTaskDetails>
    {
        let scopes = ['Group.ReadWrite.All'];
        let taskDetails = await this.getJson(`/planner/tasks/${taskId}/details`, scopes) as {value: MicrosoftGraph.PlannerTaskDetails};
        return taskDetails.value;
    }

    public async setTaskDetails(taskId: string, newData: MicrosoftGraph.PlannerTask = null): Promise<any>
    {
        let scopes = ['Group.ReadWrite.All'];
        let result = await this.patch(`/planner/tasks/${taskId}`, scopes, newData);

        return result;
    }

    public async setTaskComplete(taskId: string)
    {
        return await this.setTaskDetails(taskId, {
            percentComplete: 100
        });
    }

    public async addTask(planId: string, newTask: MicrosoftGraph.PlannerTask)
    {
        let scopes = ['Group.ReadWrite.All'];
        let result = await this.post('/planner/tasks', scopes, {
            ...newTask,
            planId,
        });

        return result;
    }

    public async removeTask( taskId: string ){
        let scopes = ['Group.ReadWrite.All'];
        let result = await this.delete(`/planner/tasks/${taskId}`, scopes);

        return result;
    }
}
