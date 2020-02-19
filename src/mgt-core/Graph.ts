/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Client, GraphRequest, MiddlewareOptions, ResponseType } from '@microsoft/microsoft-graph-client';
import {
  Contact,
  Event,
  Person,
  PlannerBucket,
  PlannerPlan,
  PlannerTask,
  User
} from '@microsoft/microsoft-graph-types';
import { Batch, blobToBase64, ComponentMiddlewareOptions, IGraph, prepScopes } from '.';
import { BetaGraph } from './BetaGraph';

/**
 * The version of the Graph to use for making requests.
 */
const GRAPH_VERSION = '1.0';

/**
 * The base Graph implementation.
 *
 * @export
 * @abstract
 * @class BaseGraph
 */
export class Graph implements IGraph {
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

  /**
   * the version of the graph to query
   *
   * @readonly
   * @protected
   * @type {string}
   * @memberof Graph
   */
  protected get version(): string {
    return this._version;
  }

  private _client: Client;
  private _version: string;
  private _componentName: string;

  constructor(client: Client, version: string = GRAPH_VERSION) {
    this._client = client;
    this._version = version;
  }

  /**
   * create a BetaGraph instance based on this Graph instance.
   *
   * @returns {BetaGraph}
   * @memberof Graph
   */
  public beta(): BetaGraph {
    const graph = new BetaGraph(this._client);
    graph._componentName = this._componentName;
    return graph;
  }

  /**
   * Returns a new instance of the Graph using the same
   * client within the context of the provider.
   *
   * @param {Element} component
   * @returns {IGraph}
   * @memberof BaseGraph
   */
  public forComponent(component: Element): Graph {
    const graph = new Graph(this._client, this._version);
    this.setComponent(component);
    return graph;
  }

  /**
   * Returns a new graph request for a specific component
   * Used internally for analytics purposes
   *
   * @param {string} path
   * @memberof BaseGraph
   */
  public api(path: string): GraphRequest {
    let request = this._client.api(path).version(this._version);

    if (this._componentName) {
      request.middlewareOptions = (options: MiddlewareOptions[]): GraphRequest => {
        const requestObj = request as any;
        requestObj._middlewareOptions = requestObj._middlewareOptions.concat(options);
        return request;
      };
      request = request.middlewareOptions([new ComponentMiddlewareOptions(this._componentName)]);
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
   * @returns {(Promise<User>)}
   * @memberof BaseGraph
   */
  public getMe(): Promise<User> {
    return this.api('me')
      .middlewareOptions(prepScopes('user.read'))
      .get();
  }

  /**
   * async promise, returns all Graph users associated with the userPrincipleName provided
   *
   * @param {string} userPrincipleName
   * @returns {(Promise<User>)}
   * @memberof BaseGraph
   */
  public getUser(userPrincipleName: string): Promise<User> {
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
    return this.getPhotoForResource(`me/contacts/${contactId}`, ['contacts.read']);
  }

  /**
   * async promise, returns Graph photo associated with provided userId
   * @param userId
   * @returns {Promise<string>}
   * @memberof BaseGraph
   */
  public getUserPhoto(userId: string): Promise<string> {
    return this.getPhotoForResource(`users/${userId}`, ['user.readbasic.all']);
  }

  /**
   * async promise, returns Graph photo associated with the logged in user
   * @returns {Promise<string>}
   * @memberof BaseGraph
   */
  public myPhoto(): Promise<string> {
    return this.getPhotoForResource('me', ['user.read']);
  }

  ///
  /// PERSON
  ///

  /**
   * async promise, returns all Graph people who are most relevant contacts to the signed in user.
   *
   * @param {string} query
   * @returns {(Promise<Person[]>)}
   * @memberof BaseGraph
   */
  public async findPerson(query: string): Promise<Person[]> {
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
   * @returns {(Promise<Person[]>)}
   * @memberof BaseGraph
   */
  public async getPeople(): Promise<Person[]> {
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
   * @returns {(Promise<Person[]>)}
   * @memberof BaseGraph
   */
  public async getPeopleFromGroup(groupId: string): Promise<Person[]> {
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
   * @returns {(Promise<Contact[]>)}
   * @memberof BaseGraph
   */
  public async findContactByEmail(email: string): Promise<Contact[]> {
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
   * @returns {(Promise<Array<Person | Contact>>)}
   * @memberof BaseGraph
   */
  public findUserByEmail(email: string): Promise<Array<Person | Contact>> {
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
   * @returns {(Promise<Event[]>)}
   * @memberof BaseGraph
   */
  public async getEvents(startDateTime: Date, endDateTime: Date, groupId?: string): Promise<Event[]> {
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
   * @param {(PlannerTask)} newTask
   * @returns {Promise<any>}
   * @memberof BaseGraph
   */
  public addPlannerTask(newTask: PlannerTask): Promise<any> {
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
   * @returns {(Promise<PlannerPlan[]>)}
   * @memberof BaseGraph
   */
  public async getAllMyPlannerPlans(): Promise<PlannerPlan[]> {
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
   * @returns {(Promise<PlannerBucket[]>)}
   * @memberof BaseGraph
   */
  public async getBucketsForPlannerPlan(planId: string): Promise<PlannerBucket[]> {
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
   * @returns {(Promise<PlannerPlan[]>)}
   * @memberof BaseGraph
   */
  public async getPlansForGroup(groupId: string): Promise<PlannerPlan[]> {
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
   * @returns {(Promise<PlannerPlan>)}
   * @memberof BaseGraph
   */
  public async getSinglePlannerPlan(planId: string): Promise<PlannerPlan> {
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
   * @returns {(Promise<PlannerTask[][]>)}
   * @memberof BaseGraph
   */
  public async getTasksForPlannerBucket(bucketId: string): Promise<PlannerTask[][]> {
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
   * @param {(PlannerTask)} details
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof BaseGraph
   */
  public async setPlannerTaskDetails(taskId: string, details: PlannerTask, eTag: string): Promise<any> {
    return await this.api(`/planner/tasks/${taskId}`)
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes('Group.ReadWrite.All'))
      .header('If-Match', eTag)
      .patch(JSON.stringify(details));
  }

  /**
   * sets the component name used in request headers.
   *
   * @protected
   * @param {Element} component
   * @memberof Graph
   */
  protected setComponent(component: Element): void {
    this._componentName = component.tagName;
  }

  /**
   * retrieves a photo for the specified resource.
   *
   * @param {string} resource
   * @param {string[]} scopes
   * @returns {Promise<string>}
   */
  protected async getPhotoForResource(resource: string, scopes: string[]): Promise<string> {
    try {
      const blob = await this.api(`${resource}/photo/$value`)
        .responseType(ResponseType.BLOB)
        .middlewareOptions(prepScopes(...scopes))
        .get();
      return await blobToBase64(blob);
    } catch (e) {
      return null;
    }
  }
}
