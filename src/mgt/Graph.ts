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
  RetryHandler,
  RetryHandlerOptions,
  TelemetryHandler
} from '@microsoft/microsoft-graph-client';
import {
  Contact,
  Event,
  Person,
  PlannerBucket,
  PlannerPlan,
  PlannerTask,
  User
} from '@microsoft/microsoft-graph-types';
import { BaseGraph, chainMiddleware, IProvider, SdkVersionMiddleware } from '../mgt-core';
import { PACKAGE_VERSION } from '../version';

/**
 * The version of the Graph to use for making requests.
 */
const GRAPH_VERSION = 'v1.0';

/**
 * Creates async methods for requesting data from the Graph
 *
 * @export
 * @class Graph
 */
export class Graph extends BaseGraph {
  constructor(providerOrClient: IProvider | Client) {
    const client = providerOrClient as Client;
    const provider = providerOrClient as IProvider;

    if (client.api) {
      super(client, GRAPH_VERSION);
    } else if (provider) {
      const middleware: Middleware[] = [
        new AuthenticationHandler(provider),
        new RetryHandler(new RetryHandlerOptions()),
        new TelemetryHandler(),
        new SdkVersionMiddleware(PACKAGE_VERSION),
        new HTTPMessageHandler()
      ];

      super(
        Client.initWithMiddleware({
          middleware: chainMiddleware(...middleware)
        }),
        GRAPH_VERSION
      );
    }
  }

  /**
   * Returns a new instance of the Graph using the same
   * client within the context of the provider.
   *
   * @param {Element} component
   * @returns {Graph}
   * @memberof Graph
   */
  public forComponent(component: Element): Graph {
    const graph = new Graph(this.client);
    graph.componentName = component.tagName.toLowerCase();
    return graph;
  }

  ///
  /// USER
  ///

  /**
   *  async promise, returns Graph User data relating to the user logged in
   *
   * @returns {Promise<User>}
   * @memberof Graph
   */
  public getMe(): Promise<User> {
    return super.getMe() as Promise<User>;
  }

  /**
   * async promise, returns all Graph users associated with the userPrincipleName provided
   *
   * @param {string} userPrincipleName
   * @returns {Promise<User>}
   * @memberof Graph
   */
  public getUser(userPrincipleName: string): Promise<User> {
    return super.getUser(userPrincipleName) as Promise<User>;
  }

  ///
  /// PERSON
  ///

  /**
   * async promise, returns all Graph people who are most relevant contacts to the signed in user.
   *
   * @param {string} query
   * @returns {Promise<Person[]>}
   * @memberof Graph
   */
  public findPerson(query: string): Promise<Person[]> {
    return super.findPerson(query) as Promise<Person[]>;
  }

  /**
   * async promise to the Graph for People, by default, it will request the most frequent contacts for the signed in user.
   *
   * @returns {Promise<Person[]>}
   * @memberof Graph
   */
  public async getPeople(): Promise<Person[]> {
    return super.getPeople() as Promise<Person[]>;
  }

  /**
   * async promise to the Graph for People, defined by a group id
   *
   * @param {string} groupId
   * @returns {Promise<Person[]>}
   * @memberof Graph
   */
  public async getPeopleFromGroup(groupId: string): Promise<Person[]> {
    return super.getPeopleFromGroup(groupId) as Promise<Person[]>;
  }

  ///
  /// CONTACTS
  ///

  /**
   * async promise, returns a Graph contact associated with the email provided
   *
   * @param {string} email
   * @returns {Promise<Contact[]>}
   * @memberof Graph
   */
  public findContactByEmail(email: string): Promise<Contact[]> {
    return super.findContactByEmail(email) as Promise<Contact[]>;
  }

  /**
   * async promise, returns Graph contact and/or Person associated with the email provided
   * Uses: Graph.findPerson(email) and Graph.findContactByEmail(email)
   *
   * @param {string} email
   * @returns {(Promise<Array<Person | Contact>>)}
   * @memberof Graph
   */
  public async findUserByEmail(email: string): Promise<Array<Person | Contact>> {
    return super.findUserByEmail(email) as Promise<Array<Person | Contact>>;
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
   * @returns {Promise<Event[]>}
   * @memberof Graph
   */
  public async getEvents(startDateTime: Date, endDateTime: Date, groupId?: string): Promise<Event[]> {
    return super.getEvents(startDateTime, endDateTime, groupId) as Promise<Event[]>;
  }

  ///
  /// PLANNER
  ///

  /**
   *  async promise, returns all planner plans associated with the user logged in
   *
   * @returns {Promise<PlannerPlan[]>}
   * @memberof Graph
   */
  public async getAllMyPlannerPlans(): Promise<PlannerPlan[]> {
    return super.getAllMyPlannerPlans() as Promise<PlannerPlan[]>;
  }

  /**
   * async promise, returns bucket (for tasks) associated with a planId
   *
   * @param {string} planId
   * @returns {Promise<PlannerBucket[]>}
   * @memberof Graph
   */
  public async getBucketsForPlannerPlan(planId: string): Promise<PlannerBucket[]> {
    return super.getBucketsForPlannerPlan(planId) as Promise<PlannerBucket[]>;
  }

  /**
   * async promise, returns all planner plans associated with the group id
   *
   * @param {string} groupId
   * @returns {Promise<PlannerPlan[]>}
   * @memberof Graph
   */
  public async getPlansForGroup(groupId: string): Promise<PlannerPlan[]> {
    return super.getPlansForGroup(groupId) as Promise<PlannerPlan[]>;
  }

  /**
   * async promise, returns a single plan from the Graph associated with the planId
   *
   * @param {string} planId
   * @returns {Promise<PlannerPlan>}
   * @memberof Graph
   */
  public async getSinglePlannerPlan(planId: string): Promise<PlannerPlan> {
    return super.getSinglePlannerPlan(planId) as Promise<PlannerPlan>;
  }

  /**
   * async promise, returns all tasks from planner associated with a bucketId
   *
   * @param {string} bucketId
   * @returns {Promise<PlannerTask[]>}
   * @memberof Graph
   */
  public async getTasksForPlannerBucket(bucketId: string): Promise<PlannerTask[]> {
    return super.getTasksForPlannerBucket(bucketId) as Promise<PlannerTask[]>;
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
  public addTodoTask(newTask: any): Promise<any> {
    throw new Error('Beta method not implemented in Graph class.');
  }

  /**
   * async promise, returns all Outlook taskGroups associated with the logged in user
   *
   * @returns {Promise<BetaTypes.OutlookTaskGroup[]>}
   * @memberof BaseGraph
   */
  public getAllMyTodoGroups(): Promise<any> {
    throw new Error('Beta method not implemented in Graph class.');
  }

  /**
   * async promise, returns all Outlook tasks associated with a taskFolder with folderId
   *
   * @param {string} folderId
   * @returns {Promise<BetaTypes.OutlookTask[]>}
   * @memberof BaseGraph
   */
  public getAllTodoTasksForFolder(folderId: string): Promise<any> {
    throw new Error('Beta method not implemented in Graph class.');
  }

  /**
   * async promise, returns all Outlook taskFolders associated with groupId
   *
   * @param {string} groupId
   * @returns {Promise<BetaTypes.OutlookTaskFolder[]>}
   * @memberof BaseGraph
   */
  public getFoldersForTodoGroup(groupId: string): Promise<any> {
    throw new Error('Beta method not implemented in Graph class.');
  }

  /**
   * async promise, returns to-do tasks from Outlook groups associated with a groupId
   *
   * @param {string} groupId
   * @returns {Promise<BetaTypes.OutlookTaskGroup>}
   * @memberof BaseGraph
   */
  public getSingleTodoGroup(groupId: string): Promise<any> {
    throw new Error('Beta method not implemented in Graph class.');
  }

  /**
   * async promise, allows developer to remove task based on taskId
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof BaseGraph
   */
  public removeTodoTask(taskId: string, eTag: string): Promise<any> {
    throw new Error('Beta method not implemented in Graph class.');
  }

  /**
   * async promise, allows developer to set to-do task to completed state
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<BetaTypes.OutlookTask>}
   * @memberof BaseGraph
   */
  public setTodoTaskComplete(taskId: string, eTag: string): Promise<any> {
    throw new Error('Beta method not implemented in Graph class.');
  }

  /**
   * async promise, allows developer to set to-do task to incomplete state
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<BetaTypes.OutlookTask>}
   * @memberof BaseGraph
   */
  public setTodoTaskIncomplete(taskId: string, eTag: string): Promise<any> {
    throw new Error('Beta method not implemented in Graph class.');
  }

  /**
   * async promise, allows developer to redefine to-do Task details associated with a taskId
   *
   * @param {string} taskId
   * @param {*} task
   * @param {string} eTag
   * @returns {Promise<BetaTypes.OutlookTask>}
   * @memberof BaseGraph
   */
  public setTodoTaskDetails(taskId: string, task: any, eTag: string): Promise<any> {
    throw new Error('Beta method not implemented in Graph class.');
  }
}
