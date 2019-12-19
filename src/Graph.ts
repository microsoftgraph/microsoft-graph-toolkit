/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  AuthenticationHandler,
  Client,
  HTTPMessageHandler,
  Middleware,
  ResponseType,
  RetryHandler,
  RetryHandlerOptions,
  TelemetryHandler
} from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import { MgtBaseComponent } from './components/baseComponent';
import { Providers } from './Providers';
import { IProvider } from './providers/IProvider';
import { Batch } from './utils/Batch';
import { CustomHeaderMiddleware } from './utils/CustomHeaderMiddleware';
import { prepScopes } from './utils/GraphHelpers';
import { SdkVersionMiddleware } from './utils/SdkVersionMiddleware';

/**
 * Creates async methods for requesting data from the Graph
 *
 * @export
 * @class Graph
 */
export class Graph {
  /**
   * middleware authentication handler
   *
   * @type {Client}
   * @memberof Graph
   */
  public client: Client;

  private _provider?: IProvider;
  private _component?: MgtBaseComponent;

  constructor(provider: IProvider, component?: MgtBaseComponent) {
    if (provider) {
      this._provider = provider;

      const middleware: Middleware[] = [
        new AuthenticationHandler(provider),
        new RetryHandler(new RetryHandlerOptions()),
        new TelemetryHandler(),
        new SdkVersionMiddleware()
      ];

      if (component) {
        this._component = component;
        middleware.push(new CustomHeaderMiddleware(this.addComponentHeader.bind(this)));
      }

      middleware.push(new HTTPMessageHandler());

      this.client = Client.initWithMiddleware({
        middleware: this.chainMiddleware(...middleware)
      });
    }
  }

  /**
   * Returns a new instance of the Graph using the same
   * provider and the provided component.
   *
   * @param {MgtBaseComponent} component
   * @returns
   * @memberof Graph
   */
  public forComponent(component: MgtBaseComponent): Graph {
    return new Graph(this._provider, component);
  }

  /**
   * creates batch request
   *
   * @returns
   * @memberof Graph
   */
  public createBatch() {
    return new Batch(this.client);
  }

  /**
   *  async promise, returns Graph User data relating to the user logged in
   *
   * @returns {Promise<MicrosoftGraph.User>}
   * @memberof Graph
   */
  public async getMe(): Promise<MicrosoftGraph.User> {
    return this.client
      .api('me')
      .middlewareOptions(prepScopes('user.read'))
      .get();
  }

  /**
   * async promise, returns all Graph users associated with the userPrincipleName provided
   *
   * @param {string} userPrincipleName
   * @returns {Promise<MicrosoftGraph.User>}
   * @memberof Graph
   */
  public async getUser(userPrincipleName: string): Promise<MicrosoftGraph.User> {
    const scopes = 'user.readbasic.all';
    return this.client
      .api(`/users/${userPrincipleName}`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  }

  /**
   * async promise, returns all Graph people who are most relevant contacts to the signed in user.
   *
   * @param {string} query
   * @returns {Promise<MicrosoftGraph.Person[]>}
   * @memberof Graph
   */
  public async findPerson(query: string): Promise<MicrosoftGraph.Person[]> {
    const scopes = 'people.read';
    const result = await this.client
      .api('/me/people')
      .search('"' + query + '"')
      .middlewareOptions(prepScopes(scopes))
      .get();
    return result ? result.value : null;
  }

  /**
   * async promise, returns a Graph contact associated with the email provided
   *
   * @param {string} email
   * @returns {Promise<MicrosoftGraph.Contact[]>}
   * @memberof Graph
   */
  public async findContactByEmail(email: string): Promise<MicrosoftGraph.Contact[]> {
    const scopes = 'contacts.read';
    const result = await this.client
      .api('/me/contacts')
      .filter(`emailAddresses/any(a:a/address eq '${email}')`)
      .middlewareOptions(prepScopes(scopes))
      .get();
    return result ? result.value : null;
  }

  /**
   * async promise, returns Graph contact and/or Person associated with the email provided
   * Uses: Graph.findPerson(email) and Graph.findContactByEmail(email)
   *
   * @param {string} email
   * @returns {(Promise<Array<MicrosoftGraph.Person | MicrosoftGraph.Contact>>)}
   * @memberof Graph
   */
  public async findUserByEmail(email: string): Promise<Array<MicrosoftGraph.Person | MicrosoftGraph.Contact>> {
    return Promise.all([this.findPerson(email), this.findContactByEmail(email)]).then(([people, contacts]) => {
      return (people || []).concat(contacts || []);
    });
  }

  /**
   * async promise, returns Graph photo associated with the logged in user
   *
   * @returns {Promise<string>}
   * @memberof Graph
   */
  public async myPhoto(): Promise<string> {
    return this.getPhotoForResource('me', ['user.read']);
  }

  /**
   * async promise, returns Graph photo associated with provided userId
   *
   * @param {string} userId
   * @returns {Promise<string>}
   * @memberof Graph
   */
  public async getUserPhoto(userId: string): Promise<string> {
    return this.getPhotoForResource(`users/${userId}`, ['user.readbasic.all']);
  }

  /**
   * async promise, returns Graph photos associated with contacts of the logged in user
   *
   * @param {string} contactId
   * @returns {Promise<string>}
   * @memberof Graph
   */
  public async getContactPhoto(contactId: string): Promise<string> {
    return this.getPhotoForResource(`me/contacts/${contactId}`, ['contacts.read']);
  }

  /**
   * async promise, returns Calender events associated with either the logged in user or a specific groupId
   *
   * @param {Date} startDateTime
   * @param {Date} endDateTime
   * @param {string} [groupId]
   * @returns {Promise<MicrosoftGraph.Event[]>}
   * @memberof Graph
   */
  public async getEvents(startDateTime: Date, endDateTime: Date, groupId?: string): Promise<MicrosoftGraph.Event[]> {
    const scopes = 'calendars.read';

    const sdt = `startdatetime=${startDateTime.toISOString()}`;
    const edt = `enddatetime=${endDateTime.toISOString()}`;

    let uri: string;

    if (groupId) {
      uri = `groups/${groupId}/calendar`;
    } else {
      uri = 'me';
    }

    uri += `/calendarview?${sdt}&${edt}`;

    const calendarView = await this.client
      .api(uri)
      .middlewareOptions(prepScopes(scopes))
      .orderby('start/dateTime')
      .get();
    return calendarView ? calendarView.value : null;
  }

  /**
   * async promise to the Graph for People, by default, it will request the most frequent contacts for the signed in user.
   *
   * @returns {Promise<MicrosoftGraph.Person[]>}
   * @memberof Graph
   */
  public async getPeople(): Promise<MicrosoftGraph.Person[]> {
    const scopes = 'people.read';

    const uri = '/me/people';
    const people = await this.client
      .api(uri)
      .middlewareOptions(prepScopes(scopes))
      .filter("personType/class eq 'Person'")
      .get();
    return people ? people.value : null;
  }

  /**
   * async promise to the Graph for People, defined by a group id
   *
   * @param {string} groupId
   * @returns {Promise<MicrosoftGraph.Person[]>}
   * @memberof Graph
   */
  public async getPeopleFromGroup(groupId: string): Promise<MicrosoftGraph.Person[]> {
    const scopes = 'people.read';

    const uri = `/groups/${groupId}/members`;
    const people = await this.client
      .api(uri)
      .middlewareOptions(prepScopes(scopes))
      .get();
    return people ? people.value : null;
  }

  /**
   * async promise, returns all planner plans associated with the group id
   *
   * @param {string} groupId
   * @returns {Promise<MicrosoftGraph.PlannerPlan[]>}
   * @memberof Graph
   */
  public async getPlansForGroup(groupId: string): Promise<MicrosoftGraph.PlannerPlan[]> {
    const scopes = 'Group.Read.All';

    const uri = `/groups/${groupId}/planner/plans`;
    const plans = await this.client
      .api(uri)
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes(scopes))
      .get();
    return plans ? plans.value : null;
  }

  /**
   *  async promise, returns all planner plans associated with the user logged in
   *
   * @returns {Promise<MicrosoftGraph.PlannerPlan[]>}
   * @memberof Graph
   */
  public async planner_getAllMyPlans(): Promise<MicrosoftGraph.PlannerPlan[]> {
    const plans = await this.client
      .api('/me/planner/plans')
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes('Group.Read.All'))
      .get();

    return plans && plans.value;
  }

  /**
   * async promise, returns a single plan from the Graph associated with the planId
   *
   * @param {string} planId
   * @returns {Promise<MicrosoftGraph.PlannerPlan>}
   * @memberof Graph
   */
  public async planner_getSinglePlan(planId: string): Promise<MicrosoftGraph.PlannerPlan> {
    const plan = await this.client
      .api(`/planner/plans/${planId}`)
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes('Group.Read.All'))
      .get();

    return plan;
  }

  /**
   * async promise, returns bucket (for tasks) associated with a planId
   *
   * @param {string} planId
   * @returns {Promise<MicrosoftGraph.PlannerBucket[]>}
   * @memberof Graph
   */
  public async planner_getBucketsForPlan(planId: string): Promise<MicrosoftGraph.PlannerBucket[]> {
    const buckets = await this.client
      .api(`/planner/plans/${planId}/buckets`)
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes('Group.Read.All'))
      .get();

    return buckets && buckets.value;
  }

  /**
   * async promise, returns all tasks from planner associated with a bucketId
   *
   * @param {string} bucketId
   * @returns {Promise<MicrosoftGraph.PlannerTask[]>}
   * @memberof Graph
   */
  public async planner_getTasksForBucket(bucketId: string): Promise<MicrosoftGraph.PlannerTask[]> {
    const tasks = await this.client
      .api(`/planner/buckets/${bucketId}/tasks`)
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes('Group.Read.All'))
      .get();

    return tasks && tasks.value;
  }

  /**
   * async promise, allows developer to set details of planner task associated with a taskId
   *
   * @param {string} taskId
   * @param {MicrosoftGraph.PlannerTask} details
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof Graph
   */
  public async planner_setTaskDetails(taskId: string, details: MicrosoftGraph.PlannerTask, eTag: string): Promise<any> {
    return await this.client
      .api(`/planner/tasks/${taskId}`)
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes('Group.ReadWrite.All'))
      .header('If-Match', eTag)
      .patch(JSON.stringify(details));
  }

  /**
   * async promise, allows developer to set a task to complete, associated with taskId
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof Graph
   */
  public async planner_setTaskComplete(taskId: string, eTag: string): Promise<any> {
    return this.planner_setTaskDetails(
      taskId,
      {
        percentComplete: 100
      },
      eTag
    );
  }

  /**
   * async promise, allows developer to set a task to incomplete, associated with taskId
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof Graph
   */
  public async planner_setTaskIncomplete(taskId: string, eTag: string): Promise<any> {
    return this.planner_setTaskDetails(
      taskId,
      {
        percentComplete: 0
      },
      eTag
    );
  }

  /**
   * async promise, allows developer to assign people to task
   *
   * @param {string} taskId
   * @param {*} people
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof Graph
   */
  public async planner_assignPeopleToTask(taskId: string, people: any, eTag: string): Promise<any> {
    return this.planner_setTaskDetails(
      taskId,
      {
        assignments: people
      },
      eTag
    );
  }

  /**
   * async promise, allows developer to create new Planner task
   *
   * @param {MicrosoftGraph.PlannerTask} newTask
   * @returns {Promise<any>}
   * @memberof Graph
   */
  public async planner_addTask(newTask: MicrosoftGraph.PlannerTask): Promise<any> {
    return this.client
      .api('/planner/tasks')
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes('Group.ReadWrite.All'))
      .post(newTask);
  }

  /**
   * async promise, allows developer to remove Planner task associated with taskId
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof Graph
   */
  public async planner_removeTask(taskId: string, eTag: string): Promise<any> {
    return this.client
      .api(`/planner/tasks/${taskId}`)
      .header('Cache-Control', 'no-store')
      .header('If-Match', eTag)
      .middlewareOptions(prepScopes('Group.ReadWrite.All'))
      .delete();
  }

  // Todo Methods

  /**
   * async promise, returns all Outlook taskGroups associated with the logged in user
   *
   * @returns {Promise<MicrosoftGraphBeta.OutlookTaskGroup[]>}
   * @memberof Graph
   */
  public async todo_getAllMyGroups(): Promise<MicrosoftGraphBeta.OutlookTaskGroup[]> {
    const groups = await this.client
      .api('/me/outlook/taskGroups')
      .header('Cache-Control', 'no-store')
      .version('beta')
      .middlewareOptions(prepScopes('Tasks.Read'))
      .get();

    return groups && groups.value;
  }

  /**
   * async promise, returns to-do tasks from Outlook groups associated with a groupId
   *
   * @param {string} groupId
   * @returns {Promise<MicrosoftGraphBeta.OutlookTaskGroup>}
   * @memberof Graph
   */
  public async todo_getSingleGroup(groupId: string): Promise<MicrosoftGraphBeta.OutlookTaskGroup> {
    const group = await this.client
      .api(`/me/outlook/taskGroups/${groupId}`)
      .header('Cache-Control', 'no-store')
      .version('beta')
      .middlewareOptions(prepScopes('Tasks.Read'))
      .get();

    return group;
  }

  /**
   * async promise, returns all Outlook taskFolders associated with groupId
   *
   * @param {string} groupId
   * @returns {Promise<MicrosoftGraphBeta.OutlookTaskFolder[]>}
   * @memberof Graph
   */
  public async todo_getFoldersForGroup(groupId: string): Promise<MicrosoftGraphBeta.OutlookTaskFolder[]> {
    const folders = await this.client
      .api(`/me/outlook/taskGroups/${groupId}/taskFolders`)
      .header('Cache-Control', 'no-store')
      .version('beta')
      .middlewareOptions(prepScopes('Tasks.Read'))
      .get();

    return folders && folders.value;
  }

  /**
   * async promise, returns all Outlook tasks associated with a taskFolder with folderId
   *
   * @param {string} folderId
   * @returns {Promise<MicrosoftGraphBeta.OutlookTask[]>}
   * @memberof Graph
   */
  public async todo_getAllTasksForFolder(folderId: string): Promise<MicrosoftGraphBeta.OutlookTask[]> {
    const tasks = await this.client
      .api(`/me/outlook/taskFolders/${folderId}/tasks`)
      .header('Cache-Control', 'no-store')
      .version('beta')
      .middlewareOptions(prepScopes('Tasks.Read'))
      .get();

    return tasks && tasks.value;
  }

  /**
   * async promise, allows developer to redefine to-do Task details associated with a taskId
   *
   * @param {string} taskId
   * @param {*} task
   * @param {string} eTag
   * @returns {Promise<MicrosoftGraphBeta.OutlookTask>}
   * @memberof Graph
   */
  public async todo_setTaskDetails(taskId: string, task: any, eTag: string): Promise<MicrosoftGraphBeta.OutlookTask> {
    return await this.client
      .api(`/me/outlook/tasks/${taskId}`)
      .header('Cache-Control', 'no-store')
      .version('beta')
      .header('If-Match', eTag)
      .middlewareOptions(prepScopes('Tasks.ReadWrite'))
      .patch(task);
  }

  /**
   * async promise, allows developer to set to-do task to completed state
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<MicrosoftGraphBeta.OutlookTask>}
   * @memberof Graph
   */
  public async todo_setTaskComplete(taskId: string, eTag: string): Promise<MicrosoftGraphBeta.OutlookTask> {
    return await this.todo_setTaskDetails(
      taskId,
      {
        isReminderOn: false,
        status: 'completed'
      },
      eTag
    );
  }

  /**
   * async promise, allows developer to set to-do task to incomplete state
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<MicrosoftGraphBeta.OutlookTask>}
   * @memberof Graph
   */
  public async todo_setTaskIncomplete(taskId: string, eTag: string): Promise<MicrosoftGraphBeta.OutlookTask> {
    return await this.todo_setTaskDetails(
      taskId,
      {
        isReminderOn: true,
        status: 'notStarted'
      },
      eTag
    );
  }

  /**
   * async promise, allows developer to add new to-do task
   *
   * @param {*} newTask
   * @returns {Promise<MicrosoftGraphBeta.OutlookTask>}
   * @memberof Graph
   */
  public async todo_addTask(newTask: any): Promise<MicrosoftGraphBeta.OutlookTask> {
    const { parentFolderId = null } = newTask;

    if (parentFolderId) {
      return await this.client
        .api(`/me/outlook/taskFolders/${parentFolderId}/tasks`)
        .header('Cache-Control', 'no-store')
        .version('beta')
        .middlewareOptions(prepScopes('Tasks.ReadWrite'))
        .post(newTask);
    } else {
      return await this.client
        .api('/me/outlook/tasks')
        .header('Cache-Control', 'no-store')
        .version('beta')
        .middlewareOptions(prepScopes('Tasks.ReadWrite'))
        .post(newTask);
    }
  }

  /**
   * async promise, allows developer to remove task based on taskId
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof Graph
   */
  public async todo_removeTask(taskId: string, eTag: string): Promise<any> {
    return await this.client
      .api(`/me/outlook/tasks/${taskId}`)
      .header('Cache-Control', 'no-store')
      .version('beta')
      .header('If-Match', eTag)
      .middlewareOptions(prepScopes('Tasks.ReadWrite'))
      .delete();
  }

  private chainMiddleware(...middlewares: Middleware[]): Middleware {
    const rootMiddleware = middlewares[0];
    let current = rootMiddleware;
    for (let i = 1; i < middlewares.length; ++i) {
      const next = middlewares[i];
      if (current.setNext) {
        current.setNext(next);
      }
      current = next;
    }
    return rootMiddleware;
  }

  private addComponentHeader(): Promise<object> {
    return new Promise((resolve, reject) => {
      resolve(this._component ? { component: this._component.tagName } : null);
    });
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

  private async getPhotoForResource(resource: string, scopes: string[]): Promise<string> {
    try {
      const blob = await this.client
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
}
