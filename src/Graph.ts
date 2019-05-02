import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import { Client } from '@microsoft/microsoft-graph-client/lib/es/Client';
import { IProvider } from './providers/IProvider';
import { ResponseType } from '@microsoft/microsoft-graph-client/lib/es/ResponseType';
import { AuthenticationHandlerOptions } from '@microsoft/microsoft-graph-client/lib/es/middleware/options/AuthenticationHandlerOptions';

export function prepScopes(...scopes: string[]) {
  const authProviderOptions = {
    scopes: scopes
  };
  return [new AuthenticationHandlerOptions(undefined, authProviderOptions)];
}

export class Graph {
  public client: Client;

  constructor(provider: IProvider) {
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
    return this.client
      .api('me')
      .middlewareOptions(prepScopes('user.read'))
      .get();
  }

  async getUser(userPrincipleName: string): Promise<MicrosoftGraph.User> {
    let scopes = 'user.readbasic.all';
    return this.client
      .api(`/users/${userPrincipleName}`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  }

  async findPerson(query: string): Promise<MicrosoftGraph.Person[]> {
    let scopes = 'people.read';
    let result = await this.client
      .api(`/me/people`)
      .search('"' + query + '"')
      .middlewareOptions(prepScopes(scopes))
      .get();
    return result ? result.value : null;
  }

  async findContactByEmail(email: string): Promise<MicrosoftGraph.Contact[]> {
    let scopes = 'contacts.read';
    let result = await this.client
      .api(`/me/contacts`)
      .filter(`emailAddresses/any(a:a/address eq '${email}')`)
      .middlewareOptions(prepScopes(scopes))
      .get();
    return result ? result.value : null;
  }

  private async getPhotoForResource(resource: string, scopes: string[]): Promise<string> {
    try {
      let blob = await this.client
        .api(`${resource}/photo/$value`)
        .version('beta')
        .responseType(ResponseType.BLOB)
        .middlewareOptions(prepScopes(...scopes))
        .get();
      return await this.blobToBase64(blob);
    } catch (e) {
      return null;
    }
  }

  async myPhoto(): Promise<string> {
    return this.getPhotoForResource('me', ['user.read']);
  }

  async getUserPhoto(userId: string): Promise<string> {
    return this.getPhotoForResource(`users/${userId}`, ['user.readbasic.all']);
  }

  async getContactPhoto(contactId: string): Promise<string> {
    return this.getPhotoForResource(`me/contacts/${contactId}`, ['contacts.read']);
  }

  async getEvents(startDateTime: Date, endDateTime: Date): Promise<Array<MicrosoftGraph.Event>> {
    let scopes = 'calendars.read';

    let sdt = `startdatetime=${startDateTime.toISOString()}`;
    let edt = `enddatetime=${endDateTime.toISOString()}`;
    let uri = `/me/calendarview?${sdt}&${edt}`;

    let calendarView = await this.client
      .api(uri)
      .middlewareOptions(prepScopes(scopes))
      .orderby('start/dateTime')
      .get();
    return calendarView ? calendarView.value : null;
  }

  // Planner Methods
  public async planner_getAllMyPlans(): Promise<MicrosoftGraph.PlannerPlan[]> {
    let plans = await this.client
      .api('/me/planner/plans')
      .header("Cache-Control", "no-store")
      .middlewareOptions(prepScopes('Group.Read.All'))
      .get();

    return plans && plans.value;
  }
  public async planner_getSinglePlan(planId: string): Promise<MicrosoftGraph.PlannerPlan> {
    let plan = await this.client
      .api(`/planner/plans/${planId}`)
      .header("Cache-Control", "no-store")
      .middlewareOptions(prepScopes('Group.Read.All'))
      .get();

    return plan;
  }
  public async planner_getBucketsForPlan(planId: string): Promise<MicrosoftGraph.PlannerBucket[]> {
    let buckets = await this.client
      .api(`/planner/plans/${planId}/buckets`)
      .header("Cache-Control", "no-store")
      .middlewareOptions(prepScopes('Group.Read.All'))
      .get();

    return buckets && buckets.value;
  }
  public async planner_getTasksForBucket(bucketId: string): Promise<MicrosoftGraph.PlannerTask[]> {
    let tasks = await this.client
      .api(`/planner/buckets/${bucketId}/tasks`)
      .header("Cache-Control", "no-store")
      .middlewareOptions(prepScopes('Group.Read.All'))
      .get();

    return tasks && tasks.value;
  }
  public async planner_setTaskDetails(taskId: string, details: MicrosoftGraph.PlannerTask, eTag: string): Promise<any> {
    return await this.client
      .api(`/planner/tasks/${taskId}`)
      .header("Cache-Control", "no-store")
      .middlewareOptions(prepScopes('Group.ReadWrite.All'))
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
    return this.client
      .api(`/planner/tasks`)
      .header("Cache-Control", "no-store")
      .middlewareOptions(prepScopes('Group.ReadWrite.All'))
      .post(newTask);
  }
  public async planner_removeTask(taskId: string, eTag: string): Promise<any> {
    return this.client
      .api(`/planner/tasks/${taskId}`)
      .header("Cache-Control", "no-store")
      .header('If-Match', eTag)
      .middlewareOptions(prepScopes('Group.ReadWrite.All'))
      .delete();
  }

  // Todo Methods
  public async todo_getAllMyGroups(): Promise<MicrosoftGraphBeta.OutlookTaskGroup[]> {
    let groups = await this.client
      .api('/me/outlook/taskGroups')
      .header("Cache-Control", "no-store")
      .version('beta')
      .middlewareOptions(prepScopes('Tasks.ReadWrite'))
      .get();

    return groups && groups.value;
  }
  public async todo_getSingleGroup(groupId: string): Promise<MicrosoftGraphBeta.OutlookTaskGroup> {
    let group = await this.client
      .api(`/me/outlook/taskGroups/${groupId}`)
      .header("Cache-Control", "no-store")
      .version('beta')
      .middlewareOptions(prepScopes('Tasks.ReadWrite'))
      .get();

    return group;
  }
  public async todo_getFoldersForGroup(groupId: string): Promise<MicrosoftGraphBeta.OutlookTaskFolder[]> {
    let folders = await this.client
      .api(`/me/outlook/taskGroups/${groupId}/taskFolders`)
      .header("Cache-Control", "no-store")
      .version('beta')
      .middlewareOptions(prepScopes('Tasks.ReadWrite'))
      .get();

    return folders && folders.value;
  }
  public async todo_getAllTasksForFolder(folderId: string): Promise<MicrosoftGraphBeta.OutlookTask[]> {
    let tasks = await this.client
      .api(`/me/outlook/taskFolders/${folderId}/tasks`)
      .header("Cache-Control", "no-store")
      .version('beta')
      .middlewareOptions(prepScopes('Tasks.ReadWrite'))
      .get();

    return tasks && tasks.value;
  }
  public async todo_setTaskDetails(taskId: string, task: any, eTag: string): Promise<MicrosoftGraphBeta.OutlookTask> {
    return await this.client
      .api(`/me/outlook/tasks/${taskId}`)
      .header("Cache-Control", "no-store")
      .version('beta')
      .header('If-Match', eTag)
      .middlewareOptions(prepScopes('Tasks.ReadWrite'))
      .patch(task);
  }
  public async todo_setTaskComplete(taskId: string, eTag: string): Promise<MicrosoftGraphBeta.OutlookTask> {
    return await this.todo_setTaskDetails(
      taskId,
      {
        status: 'completed',
        isReminderOn: false
      },
      eTag
    );
  }
  public async todo_setTaskIncomplete(taskId: string, eTag: string): Promise<MicrosoftGraphBeta.OutlookTask> {
    return await this.todo_setTaskDetails(
      taskId,
      {
        status: 'notStarted',
        isReminderOn: true
      },
      eTag
    );
  }

  public async todo_addTask(newTask: any): Promise<MicrosoftGraphBeta.OutlookTask> {
    let { parentFolderId = null } = newTask;

    if (parentFolderId)
      return await this.client
        .api(`/me/outlook/taskFolders/${parentFolderId}/tasks`)
        .header("Cache-Control", "no-store")
        .version('beta')
        .middlewareOptions(prepScopes('Tasks.ReadWrite'))
        .post(newTask);
    else
      return await this.client
        .api(`/me/outlook/tasks`)
        .header("Cache-Control", "no-store")
        .version('beta')
        .middlewareOptions(prepScopes('Tasks.ReadWrite'))
        .post(newTask);
  }

  public async todo_removeTask(taskId: string, eTag: string): Promise<any> {
    return await this.client
      .api(`/me/outlook/tasks/${taskId}`)
      .header("Cache-Control", "no-store")
      .version('beta')
      .header('If-Match', eTag)
      .middlewareOptions(prepScopes('Tasks.ReadWrite'))
      .delete();
  }
}
