/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Client, GraphRequest, MiddlewareOptions } from '@microsoft/microsoft-graph-client';
import * as GraphTypes from '@microsoft/microsoft-graph-types';
import * as BetaTypes from '@microsoft/microsoft-graph-types-beta';
import { Batch, ComponentMiddlewareOptions, getPhotoForResource, IGraph, prepScopes } from '.';

/**
 * The base Graph implementation.
 *
 * @export
 * @abstract
 * @class BaseGraph
 */
export abstract class BaseGraph implements IGraph {
  /**
   * The name of the component making the request
   *
   * @protected
   * @type {string}
   * @memberof BaseGraph
   */
  protected componentName: string;

  /**
   * the internal client used to make graph calls
   *
   * @readonly
   * @protected
   * @type {Client}
   * @memberof BaseGraph
   */
  protected get client(): Client {
    return this._client;
  }
  private _client: Client;
  private _version: string;

  constructor(client: Client, version: string) {
    this._client = client;
    this._version = version;
  }

  /**
   * Returns a new instance of the Graph using the same
   * client within the context of the provider.
   *
   * @param {Element} component
   * @returns {IGraph}
   * @memberof BaseGraph
   */
  public abstract forComponent(component: Element): IGraph;

  /**
   * Returns a new graph request for a specific component
   * Used internally for analytics purposes
   *
   * @param {string} path
   * @memberof BaseGraph
   */
  public api(path: string): GraphRequest {
    let request = this._client.api(path).version(this._version);

    if (this.componentName) {
      request.middlewareOptions = (options: MiddlewareOptions[]): GraphRequest => {
        const requestObj = request as any;
        requestObj._middlewareOptions = requestObj._middlewareOptions.concat(options);
        return request;
      };
      request = request.middlewareOptions([new ComponentMiddlewareOptions(this.componentName)]);
    }

    return request;
  }

  ///
  /// BATCH
  ///

  /**
   * creates a new batch request
   *
   * @returns {Batch}
   * @memberof BaseGraph
   */
  public createBatch(): Batch {
    return new Batch(this);
  }

  ///
  /// USER
  ///

  /**
   * async promise, returns Graph User data relating to the user logged in
   *
   * @returns {(Promise<GraphTypes.User | BetaTypes.User>)}
   * @memberof BaseGraph
   */
  public getMe(): Promise<GraphTypes.User | BetaTypes.User> {
    return this.api('me')
      .middlewareOptions(prepScopes('user.read'))
      .get();
  }

  /**
   * async promise, returns all Graph users associated with the userPrincipleName provided
   *
   * @param {string} userPrincipleName
   * @returns {(Promise<GraphTypes.User | BetaTypes.User>)}
   * @memberof BaseGraph
   */
  public getUser(userPrincipleName: string): Promise<GraphTypes.User | BetaTypes.User> {
    const scopes = 'user.readbasic.all';
    return this.api(`/users/${userPrincipleName}`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  }

  ///
  /// PHOTO
  ///

  /**
   * async promise, returns Graph photos associated with contacts of the logged in user
   * @param contactId
   * @returns {Promise<string>}
   * @memberof BaseGraph
   */
  public getContactPhoto(contactId: string): Promise<string> {
    return getPhotoForResource(`me/contacts/${contactId}`, ['contacts.read']);
  }

  /**
   * async promise, returns Graph photo associated with provided userId
   * @param userId
   * @returns {Promise<string>}
   * @memberof BaseGraph
   */
  public getUserPhoto(userId: string): Promise<string> {
    return getPhotoForResource(`users/${userId}`, ['user.readbasic.all']);
  }

  /**
   * async promise, returns Graph photo associated with the logged in user
   * @returns {Promise<string>}
   * @memberof BaseGraph
   */
  public myPhoto(): Promise<string> {
    return getPhotoForResource('me', ['user.read']);
  }

  ///
  /// PERSON
  ///

  /**
   * async promise, returns all Graph people who are most relevant contacts to the signed in user.
   *
   * @param {string} query
   * @returns {(Promise<GraphTypes.Person[] | BetaTypes.Person[]>)}
   * @memberof BaseGraph
   */
  public async findPerson(query: string): Promise<GraphTypes.Person[] | BetaTypes.Person[]> {
    const scopes = 'people.read';
    const result = await this.api('/me/people')
      .search('"' + query + '"')
      .middlewareOptions(prepScopes(scopes))
      .get();
    return result ? result.value : null;
  }

  /**
   * async promise to the Graph for People, by default, it will request the most frequent contacts for the signed in user.
   *
   * @returns {(Promise<GraphTypes.Person[] | BetaTypes.Person[]>)}
   * @memberof BaseGraph
   */
  public async getPeople(): Promise<GraphTypes.Person[] | BetaTypes.Person[]> {
    const scopes = 'people.read';

    const uri = '/me/people';
    const people = await this.api(uri)
      .middlewareOptions(prepScopes(scopes))
      .filter("personType/class eq 'Person'")
      .get();
    return people ? people.value : null;
  }

  /**
   * async promise to the Graph for People, defined by a group id
   *
   * @param {string} groupId
   * @returns {(Promise<GraphTypes.Person[] | BetaTypes.Person[]>)}
   * @memberof BaseGraph
   */
  public async getPeopleFromGroup(groupId: string): Promise<GraphTypes.Person[] | BetaTypes.Person[]> {
    const scopes = 'people.read';

    const uri = `/groups/${groupId}/members`;
    const people = await this.api(uri)
      .middlewareOptions(prepScopes(scopes))
      .get();
    return people ? people.value : null;
  }

  ///
  /// CONTACTS
  ///

  /**
   * async promise, returns a Graph contact associated with the email provided
   *
   * @param {string} email
   * @returns {(Promise<GraphTypes.Contact[] | BetaTypes.Contact[]>)}
   * @memberof BaseGraph
   */
  public async findContactByEmail(email: string): Promise<GraphTypes.Contact[] | BetaTypes.Contact[]> {
    const scopes = 'contacts.read';
    const result = await this.api('/me/contacts')
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
   * @returns {(Promise<Array<GraphTypes.Person | GraphTypes.Contact> | Array<BetaTypes.Person | BetaTypes.Contact>>)}
   * @memberof BaseGraph
   */
  public findUserByEmail(
    email: string
  ): Promise<Array<GraphTypes.Person | GraphTypes.Contact> | Array<BetaTypes.Person | BetaTypes.Contact>> {
    return Promise.all([this.findPerson(email), this.findContactByEmail(email)]).then(([people, contacts]) => {
      return ((people as any[]) || []).concat(contacts || []);
    });
  }

  ///
  /// EVENTS
  ///

  /**
   * async promise, returns Calender events associated with either the logged in user or a specific groupId
   *
   * @param {Date} startDateTime
   * @param {Date} endDateTime
   * @param {string} [groupId]
   * @returns {(Promise<GraphTypes.Event[] | BetaTypes.Event[]>)}
   * @memberof BaseGraph
   */
  public async getEvents(
    startDateTime: Date,
    endDateTime: Date,
    groupId?: string
  ): Promise<GraphTypes.Event[] | BetaTypes.Event[]> {
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

    const calendarView = await this.api(uri)
      .middlewareOptions(prepScopes(scopes))
      .orderby('start/dateTime')
      .get();
    return calendarView ? calendarView.value : null;
  }

  ///
  /// PLANNER
  ///

  /**
   * async promise, allows developer to create new Planner task
   *
   * @param {(GraphTypes.PlannerTask | BetaTypes.PlannerTask)} newTask
   * @returns {Promise<any>}
   * @memberof BaseGraph
   */
  public addPlannerTask(newTask: GraphTypes.PlannerTask | BetaTypes.PlannerTask): Promise<any> {
    return this.api('/planner/tasks')
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes('Group.ReadWrite.All'))
      .post(newTask);
  }

  /**
   * async promise, allows developer to assign people to task
   *
   * @param {string} taskId
   * @param {*} people
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof BaseGraph
   */
  public assignPeopleToPlannerTask(taskId: string, people: any, eTag: string): Promise<any> {
    return this.setPlannerTaskDetails(
      taskId,
      {
        assignments: people
      },
      eTag
    );
  }

  /**
   * async promise, returns all planner plans associated with the user logged in
   *
   * @returns {(Promise<GraphTypes.PlannerPlan[] | BetaTypes.PlannerPlan[]>)}
   * @memberof BaseGraph
   */
  public async getAllMyPlannerPlans(): Promise<GraphTypes.PlannerPlan[] | BetaTypes.PlannerPlan[]> {
    const plans = await this.api('/me/planner/plans')
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes('Group.Read.All'))
      .get();

    return plans && plans.value;
  }

  /**
   * async promise, returns bucket (for tasks) associated with a planId
   *
   * @param {string} planId
   * @returns {(Promise<GraphTypes.PlannerBucket[] | BetaTypes.PlannerBucket[]>)}
   * @memberof BaseGraph
   */
  public async getBucketsForPlannerPlan(
    planId: string
  ): Promise<GraphTypes.PlannerBucket[] | BetaTypes.PlannerBucket[]> {
    const buckets = await this.api(`/planner/plans/${planId}/buckets`)
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes('Group.Read.All'))
      .get();

    return buckets && buckets.value;
  }

  /**
   * async promise, returns all planner plans associated with the group id
   *
   * @param {string} groupId
   * @returns {(Promise<GraphTypes.PlannerPlan[] | BetaTypes.PlannerPlan[]>)}
   * @memberof BaseGraph
   */
  public async getPlansForGroup(groupId: string): Promise<GraphTypes.PlannerPlan[] | BetaTypes.PlannerPlan[]> {
    const scopes = 'Group.Read.All';

    const uri = `/groups/${groupId}/planner/plans`;
    const plans = await this.api(uri)
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes(scopes))
      .get();
    return plans ? plans.value : null;
  }

  /**
   * async promise, returns a single plan from the Graph associated with the planId
   *
   * @param {string} planId
   * @returns {(Promise<GraphTypes.PlannerPlan | BetaTypes.PlannerPlan>)}
   * @memberof BaseGraph
   */
  public async getSinglePlannerPlan(planId: string): Promise<GraphTypes.PlannerPlan | BetaTypes.PlannerPlan> {
    const plan = await this.api(`/planner/plans/${planId}`)
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes('Group.Read.All'))
      .get();

    return plan;
  }

  /**
   * async promise, returns all tasks from planner associated with a bucketId
   *
   * @param {string} bucketId
   * @returns {(Promise<GraphTypes.PlannerTask[] | BetaTypes.PlannerTask[]>)}
   * @memberof BaseGraph
   */
  public async getTasksForPlannerBucket(bucketId: string): Promise<GraphTypes.PlannerTask[] | BetaTypes.PlannerTask[]> {
    const tasks = await this.api(`/planner/buckets/${bucketId}/tasks`)
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes('Group.Read.All'))
      .get();

    return tasks && tasks.value;
  }

  /**
   * async promise, allows developer to remove Planner task associated with taskId
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof BaseGraph
   */
  public removePlannerTask(taskId: string, eTag: string): Promise<any> {
    return this.api(`/planner/tasks/${taskId}`)
      .header('Cache-Control', 'no-store')
      .header('If-Match', eTag)
      .middlewareOptions(prepScopes('Group.ReadWrite.All'))
      .delete();
  }

  /**
   * async promise, allows developer to set a task to complete, associated with taskId
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof BaseGraph
   */
  public setPlannerTaskComplete(taskId: string, eTag: string): Promise<any> {
    return this.setPlannerTaskDetails(
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
   * @memberof BaseGraph
   */
  public setPlannerTaskIncomplete(taskId: string, eTag: string): Promise<any> {
    return this.setPlannerTaskDetails(
      taskId,
      {
        percentComplete: 0
      },
      eTag
    );
  }

  /**
   * async promise, allows developer to set details of planner task associated with a taskId
   *
   * @param {string} taskId
   * @param {(GraphTypes.PlannerTask | BetaTypes.PlannerTask)} details
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof BaseGraph
   */
  public async setPlannerTaskDetails(
    taskId: string,
    details: GraphTypes.PlannerTask | BetaTypes.PlannerTask,
    eTag: string
  ): Promise<any> {
    return await this.api(`/planner/tasks/${taskId}`)
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes('Group.ReadWrite.All'))
      .header('If-Match', eTag)
      .patch(JSON.stringify(details));
  }

  ///
  /// TO-DO
  ///

  /**
   * async promise, allows developer to add new to-do task
   *
   * @param {*} newTask
   * @returns {Promise<BetaTypes.OutlookTask>}
   * @memberof BaseGraph
   */
  public abstract addTodoTask(newTask: any): Promise<BetaTypes.OutlookTask>;

  /**
   * async promise, returns all Outlook taskGroups associated with the logged in user
   *
   * @returns {Promise<BetaTypes.OutlookTaskGroup[]>}
   * @memberof BaseGraph
   */
  public abstract getAllMyTodoGroups(): Promise<BetaTypes.OutlookTaskGroup[]>;

  /**
   * async promise, returns all Outlook tasks associated with a taskFolder with folderId
   *
   * @param {string} folderId
   * @returns {Promise<BetaTypes.OutlookTask[]>}
   * @memberof BaseGraph
   */
  public abstract getAllTodoTasksForFolder(folderId: string): Promise<BetaTypes.OutlookTask[]>;

  /**
   * async promise, returns all Outlook taskFolders associated with groupId
   *
   * @param {string} groupId
   * @returns {Promise<BetaTypes.OutlookTaskFolder[]>}
   * @memberof BaseGraph
   */
  public abstract getFoldersForTodoGroup(groupId: string): Promise<BetaTypes.OutlookTaskFolder[]>;

  /**
   * async promise, returns to-do tasks from Outlook groups associated with a groupId
   *
   * @param {string} groupId
   * @returns {Promise<BetaTypes.OutlookTaskGroup>}
   * @memberof BaseGraph
   */
  public abstract getSingleTodoGroup(groupId: string): Promise<BetaTypes.OutlookTaskGroup>;

  /**
   * async promise, allows developer to remove task based on taskId
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof BaseGraph
   */
  public abstract removeTodoTask(taskId: string, eTag: string): Promise<any>;

  /**
   * async promise, allows developer to set to-do task to completed state
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<BetaTypes.OutlookTask>}
   * @memberof BaseGraph
   */
  public abstract setTodoTaskComplete(taskId: string, eTag: string): Promise<BetaTypes.OutlookTask>;

  /**
   * async promise, allows developer to set to-do task to incomplete state
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<BetaTypes.OutlookTask>}
   * @memberof BaseGraph
   */
  public abstract setTodoTaskIncomplete(taskId: string, eTag: string): Promise<BetaTypes.OutlookTask>;

  /**
   * async promise, allows developer to redefine to-do Task details associated with a taskId
   *
   * @param {string} taskId
   * @param {*} task
   * @param {string} eTag
   * @returns {Promise<BetaTypes.OutlookTask>}
   * @memberof BaseGraph
   */
  public abstract setTodoTaskDetails(taskId: string, task: any, eTag: string): Promise<BetaTypes.OutlookTask>;
}
