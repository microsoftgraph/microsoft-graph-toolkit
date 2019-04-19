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
    planner_getAllMyPlans?(): Promise<MicrosoftGraph.PlannerPlan[]>;
    planner_getSinglePlan?(planId: string): Promise<MicrosoftGraph.PlannerPlan>;
    planner_getPlanDetails?(planId: string): Promise<MicrosoftGraph.PlannerPlanDetails>;

    planner_getBucketsForPlan?(planId: string): Promise<MicrosoftGraph.PlannerBucket[]>;

    planner_getAllMyTasks?(): Promise<MicrosoftGraph.PlannerTask[]>;
    planner_getTasksForPlan?(planId: string): Promise<MicrosoftGraph.PlannerTask[]>;
    planner_getTaskDetails?(taskId: string): Promise<MicrosoftGraph.PlannerTaskDetails>;

    planner_setTaskDetails?(taskId: string, newInfo: MicrosoftGraph.PlannerTask, eTag: string): Promise<any>;
    planner_setTaskComplete?(taskId: string, eTag: string): Promise<any>;
    planner_setTaskIncomplete?(taskId: string, eTag: string): Promise<any>;
    planner_addTask?(newTask: MicrosoftGraph.PlannerTask): Promise<any>;
    planner_removeTask?(taskId: string, eTag: string): Promise<any>;

    //ToDo Task Methods
    todo_getAllMyGroups?(): Promise<any>;
    todo_getSingleGroup?(groupId: string): Promise<any>;

    todo_getFoldersForGroup?(groupId: string): Promise<any>;

    todo_getAllMyTasks?(): Promise<any>;
    todo_getAllTasksForFolder?(folderId: string): Promise<any>;

    todo_setTask?(taskId: string, task: any, eTag: string): Promise<any>;
    todo_setTaskComplete?(taskId: string, dateTime: string, eTag: string): Promise<any>;
    todo_setTaskIncomplete?(taskId: string, eTag: string): Promise<any>
    todo_addTask?(newTask: any): Promise<any>;
    todo_removeTask?(taskId: string, eTag: string): Promise<any>;
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

    async patch(resource: string, scopes: string[], data?: any, eTag?: string): Promise<Response>
    {
        resource = this.checkResource(resource);
        
        let token = await this._provider.getAccessToken(...scopes).catch(err=>null);
        if(!token)
            return null;

        let body = !!data ? {body: JSON.stringify(data)} : {};

        console.log(resource, scopes, data);

        let req = { 
            method: 'PATCH',
            headers: {
                Authorization: `Bearer ${token}`,
                "Content-Type": 'application/json',
            },
            ...body
        };

        if(eTag)
            req.headers['If-Match'] = eTag;

        let result = await fetch(this.rootUrl + resource, req);
        this.checkResponse(result);

        return result;
    }

    async post(resource: string, scopes: string[], data: any, eTag?: string)
    {
        resource = this.checkResource(resource);
        
        let token = await this._provider.getAccessToken(...scopes).catch(err=>null);
        if(!token)
            return null;

        let body = !!data ? {body: JSON.stringify(data)} : {};
        let req = {
            method: 'POST',
            headers: {
                Authorization: `Bearer ${token}`,
                "Content-Type": 'application/json',
            },
            ...body
        };

        if(eTag)
            req.headers['If-Match'] = eTag;

        let result = await fetch(this.rootUrl + resource, req);
        this.checkResponse(result);
        return result;
    }

    async delete(resource: string, scopes: string[], data?: any, eTag?: string)
    {
        resource = this.checkResource(resource);
        
        let token = await this._provider.getAccessToken(...scopes).catch(err=>null);
        if(!token)
            return null;

        let body = !!data ? {body: JSON.stringify(data)} : {};

        console.log(resource);

        let req = {
            method: 'DELETE',
            headers: {
                Authorization: `Bearer ${token}`,
                "Content-Type": 'application/json',
            },
            ...body,
        };

        if(eTag)
            req.headers['If-Match'] = eTag;

        let result = await fetch(this.rootUrl + resource, req);

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


    // Planner Task Methods
    public async planner_getAllMyPlans(): Promise<MicrosoftGraph.PlannerPlan[]>
    {
        let scopes = ['Group.ReadWrite.All'];
        let planners = await this.getJson('/me/planner/plans', scopes) as {value: MicrosoftGraph.PlannerPlan[]};
        return (planners && planners.value);
    }

    public async planner_getSinglePlan(planId: string): Promise<MicrosoftGraph.PlannerPlan>
    {
        let scopes = ['Group.ReadWrite.All'];
        let plan = await this.getJson(`/planner/plans/${planId}`, scopes) as MicrosoftGraph.PlannerPlan;

        return plan;
    }

    public async planner_getBucketsForPlan(planId: string): Promise<MicrosoftGraph.PlannerBucket[]> 
    {
        let scopes = ['Group.ReadWrite.All'];
        let buckets = await this.getJson(`/planner/plans/${planId}/buckets`, scopes) as {value: MicrosoftGraph.PlannerBucket[]};

        return (buckets && buckets.value);
    }

    public async planner_getTasksForPlan(planId: string): Promise<MicrosoftGraph.PlannerTask[]>
    {
        let scopes = ['Group.ReadWrite.All'];
        let tasks = await this.getJson(`/planner/plans/${planId}/tasks`, scopes) as {value: MicrosoftGraph.PlannerTask[]};
        return tasks.value;
    }

    public async planner_getTaskDetails(taskId: string): Promise<MicrosoftGraph.PlannerTaskDetails>
    {
        let scopes = ['Group.ReadWrite.All'];
        let taskDetails = await this.getJson(`/planner/tasks/${taskId}/details`, scopes) as {value: MicrosoftGraph.PlannerTaskDetails};
        return taskDetails.value;
    }


    public async planner_setTaskDetails(taskId: string, newData: MicrosoftGraph.PlannerTask = null, eTag: string): Promise<any>
    {
        let scopes = ['Group.ReadWrite.All'];
        let result = await this.patch(`/planner/tasks/${taskId}`, scopes, newData, eTag);

        return result;
    }

    public async planner_setTaskComplete(taskId: string, eTag: string)
    {
        return await this.planner_setTaskDetails(taskId, {
            percentComplete: 100
            },
            eTag
        );
    }

    public async planner_setTaskIncomplete(taskId: string, eTag: string)
    {
        return await this.planner_setTaskDetails( taskId, {
            percentComplete: 0
            }, 
            eTag
        );
    }

    public async planner_addTask(newTask: MicrosoftGraph.PlannerTask)
    {
        let scopes = ['Group.ReadWrite.All'];
        let result = await this.post('/planner/tasks', scopes, newTask);

        return result;
    }

    public async planner_removeTask( taskId: string, eTag: string ){
        let scopes = ['Group.ReadWrite.All'];
        let result = await this.delete(`/planner/tasks/${taskId}`, scopes, void 0, eTag);

        return result;
    }

    // Todo Task Methods

    public async todo_getAllMyGroups(){
        let scopes = ["Tasks.ReadWrite"];
        let ret = await this.getJson('/me/outlook/taskGroups', scopes) as {value: any[]};

        return (ret && ret.value);
    }
    public async todo_getSingleGroup(groupId: string){
        let scopes = ["Tasks.ReadWrite"];
        let ret = await this.getJson(`/me/outlook/taskGroups/${groupId}`, scopes);

        return (ret);
    }

    public async todo_getFoldersForGroup(groupId: string){
        let scopes = ["Tasks.ReadWrite"];
        let ret = await this.getJson(`/me/outlook/taskGroup/${groupId}/taskFolders`, scopes) as {value: any[]};

        return (ret && ret.value);
    }

    public async todo_getAllMyTasks(){
        let scopes = ["Tasks.ReadWrite"];
        let ret = await this.getJson('/me/outlook/tasks', scopes) as {value: any[]};

        return (ret && ret.value);
    }
    public async todo_getAllTasksForFolder(folderId: string){
        let scopes = ["Tasks.ReadWrite"];
        let ret = await this.getJson(`/me/outlook/taskFolders/${folderId}/tasks`, scopes) as {value: any[]};

        return (ret && ret.value);
    }

    public async todo_setTask(taskId: string, task: any, eTag: string){
        let scopes = ["Tasks.ReadWrite"];        
        let ret = await this.patch(`/me/outlook/tasks/${taskId}`, scopes, task, eTag);

        return (ret);
    }
    public async todo_setTaskComplete(taskId: string, dateTime: string, eTag: string){
        let ret = this.todo_setTask(taskId, {
            completedDateTime: {
                dateTime,
                timeZone: "UTC"
            }
        }, eTag);
        return (ret);
    }
    public async todo_setTaskIncomplete(taskId: string, eTag: string){
        let ret = this.todo_setTask(taskId, {
            completedDateTime: null
        }, eTag);
        return (ret);
    }
    public async todo_addTask(newTask: any){
        let scopes = ["Tasks.ReadWrite"];
        let ret = await this.post('/me/outlook/tasks', scopes, newTask);

        return (ret);
    }
    public async todo_removeTask(taskId: string, eTag: string){
        let scopes = ["Tasks.ReadWrite"];
        let ret = await this.delete(`/me/outlook/tasks/${taskId}`, scopes, void 0, eTag);

        return (ret);
    }
}
