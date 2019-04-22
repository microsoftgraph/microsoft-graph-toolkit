import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { Client } from '@microsoft/microsoft-graph-client/lib/es/Client';
import { IProvider } from './providers/IProvider';
import { ResponseType } from '@microsoft/microsoft-graph-client/lib/es/ResponseType';

export class Graph {
  public client: Client;

  constructor(provider: IProvider) {
    // this.token = token;
    if (provider) {
      this.client = Client.initWithMiddleware({
        authProvider: provider
      });
    }
  }

  private blobToBase64(blob: Blob): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = reject;
      reader.onload = _ => {
        resolve(reader.result as string);
      };
      reader.readAsDataURL(blob);
    });
  }

  async me(): Promise<MicrosoftGraph.User> {
    let scopes = ['user.read'];
    return this.client
      .api('me')
      .middlewareOptions([{ scopes: scopes }])
      .get();
  }

  async getUser(userPrincipleName: string): Promise<MicrosoftGraph.User> {
    let scopes = ['user.readbasic.all'];
    return this.client
      .api(`/users/${userPrincipleName}`)
      .middlewareOptions([{ scopes: scopes }])
      .get();
  }

  async findPerson(query: string): Promise<MicrosoftGraph.Person[]> {
    let scopes = ['people.read'];
    let result = await this.client
      .api(`/me/people`)
      .search('"' + query + '"')
      .middlewareOptions([{ scopes: scopes }])
      .get();
    return result ? result.value : null;
  }

  async myPhoto(): Promise<string> {
    let scopes = ['user.read'];
    let blob = await this.client
      .api('/me/photo/$value')
      .responseType(ResponseType.BLOB)
      .middlewareOptions([{ scopes: scopes }])
      .get();
    return await this.blobToBase64(blob);
  }

  async getUserPhoto(id: string): Promise<string> {
    let scopes = ['user.readbasic.all'];
    let blob = await this.client
      .api(`users/${id}/photo/$value`)
      .responseType(ResponseType.BLOB)
      .middlewareOptions([{ scopes: scopes }])
      .get();
    return await this.blobToBase64(blob);
  }

  async calendar(startDateTime: Date, endDateTime: Date): Promise<Array<MicrosoftGraph.Event>> {
    let scopes = ['calendars.read'];

    let sdt = `startdatetime=${startDateTime.toISOString()}`;
    let edt = `enddatetime=${endDateTime.toISOString()}`;
    let uri = `/me/calendarview?${sdt}&${edt}`;

    let calendarView = await this.client
      .api(uri)
      .middlewareOptions([{ scopes: scopes }])
      .get();
    return calendarView ? calendarView.value : null;
  }

  // Planner Methods
  public async planner_getAllMyPlans(): Promise<MicrosoftGraph.PlannerPlan[]> {
    let scopes = ['Group.Read.All'];

    let plans = await this.client
      .api('/me/planner/plans')
      .middlewareOptions([{ scopes }])
      .get();

    return plans && plans.value;
  }
  public async planner_getSinglePlan(planId: string): Promise<MicrosoftGraph.PlannerPlan> {
    let scopes = ['Group.Read.All'];

    let plan = await this.client
      .api(`/planner/plans/${planId}`)
      .middlewareOptions([{ scopes }])
      .get();

    return plan;
  }
  public async planner_getBucketsForPlan(planId: string): Promise<MicrosoftGraph.PlannerBucket[]> {
    let scopes = ['Group.Read.All'];

    let buckets = await this.client
      .api(`/planner/plans/${planId}/buckets`)
      .middlewareOptions([{ scopes }])
      .get();

    return buckets && buckets.value;
  }
  public async planner_getTasksForBucket(bucketId: string): Promise<MicrosoftGraph.PlannerTask[]> {
    let scopes = ['Group.Read.All'];

    let tasks = await this.client
      .api(`/planner/buckets/${bucketId}/tasks`)
      .middlewareOptions([{ scopes }])
      .get();

    return tasks && tasks.value;
  }
  public async planner_setTaskDetails(taskId: string, details: MicrosoftGraph.PlannerTask, eTag: string): Promise<any> {
    let scopes = ['Group.ReadWrite.All'];

    return await this.client
      .api(`/planner/tasks/${taskId}`)
      .middlewareOptions([{ scopes }])
      .header('If-Match', eTag)
      .patch(JSON.stringify(details));
  }
  public async planner_setTaskComplete(taskId: string, eTag: string): Promise<any> {
    return this.planner_setTaskDetails(
      taskId,
      {
        percentComplete: 100
      },
      eTag
    );
  }
  public async planner_setTaskIncomplete(taskId: string, eTag: string): Promise<any> {
    return this.planner_setTaskDetails(
      taskId,
      {
        percentComplete: 0
      },
      eTag
    );
  }
  public async planner_addTask(newTask: MicrosoftGraph.PlannerTask): Promise<any> {
    let scopes = ['Group.ReadWrite.All'];

    return this.client
      .api(`/planner/tasks`)
      .middlewareOptions([{ scopes }])
      .post(JSON.stringify(newTask));
  }
  public async planner_removeTask(taskId: string, eTag: string): Promise<any> {
    let scopes = ['Group.ReadWrite.All'];

    return this.client
      .api(`/planner/tasks/${taskId}`)
      .header('If-Match', eTag)
      .middlewareOptions([{ scopes }])
      .delete();
  }

  // Todo Methods
  public async todo_getAllMyGroups(): Promise<any> {
    let scopes = ['Tasks.Read'];

    let groups = await this.client
      .api('/me/outlook/taskGroups')
      .middlewareOptions([{ scopes }])
      .get();

    return groups && groups.value;
  }
  public async todo_getSingleGroup(groupId: string): Promise<any> {
    let scopes = ['Tasks.Read'];

    let group = await this.client
      .api(`/me/outlook/taskGroup/${groupId}`)
      .middlewareOptions([{ scopes }])
      .get();

    return group;
  }
  public async todo_getFoldersForGroup(groupId: string): Promise<any> {
    let scopes = ['Tasks.Read'];

    let folders = await this.client
      .api(`/me/outlook/taskGroup/${groupId}/taskFolders`)
      .middlewareOptions([{ scopes }])
      .get();

    return folders && folders.value;
  }
  public async todo_getAllTasksForFolder(folderId: string): Promise<any> {
    let scopes = ['Tasks.Read'];

    let tasks = await this.client
      .api(`/me/outlook/taskFolders/${folderId}/tasks`)
      .middlewareOptions([{ scopes }])
      .get();

    return tasks && tasks.value;
  }
  public async todo_setTaskDetails(taskId: string, task: any, eTag: string): Promise<any> {
    let scopes = ['Tasks.ReadWrite'];

    return await this.client
      .api(`/me/outlook/tasks/${taskId}`)
      .header('If-Match', eTag)
      .middlewareOptions([{ scopes }])
      .patch(JSON.stringify(task));
  }
  public async todo_setTaskComplete(taskId: string, dateTime: string, eTag: string): Promise<any> {
    return await this.todo_setTaskDetails(taskId, {}, eTag);
  }
  public async todo_setTaskIncomplete(taskId: string, eTag: string): Promise<any> {
    return await this.todo_setTaskDetails(taskId, {}, eTag);
  }
  public async todo_addTask(newTask: any): Promise<any> {
    let scopes = ['Tasks.ReadWrite'];

    return await this.client
      .api(`/me/outlook/tasks`)
      .middlewareOptions([{ scopes }])
      .post(JSON.stringify(newTask));
  }
  public async todo_removeTask(taskId: string, eTag: string): Promise<any> {
    let scopes = ['Tasks.ReadWrite'];

    return await this.client
      .api(`/me/outlook/tasks/${taskId}`)
      .header('If-Match', eTag)
      .middlewareOptions([{ scopes }])
      .delete();
  }
}
