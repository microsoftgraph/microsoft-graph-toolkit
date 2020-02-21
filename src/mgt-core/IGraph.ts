/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Client, GraphRequest } from '@microsoft/microsoft-graph-client';
import {
  Contact,
  Event,
  Person,
  PlannerBucket,
  PlannerPlan,
  PlannerTask,
  User
} from '@microsoft/microsoft-graph-types';
import { Batch } from './utils/Batch';

/**
 * The common functions of the Graph
 *
 * @export
 * @interface IGraph
 */
export interface IGraph {
  /**
   * the internal client used to make graph calls
   *
   * @type {Client}
   * @memberof IGraph
   */
  readonly client: Client;

  /**
   * the component name appended to Graph request headers
   *
   * @type {string}
   * @memberof IGraph
   */
  readonly componentName: string;

  /**
   * the version of the graph to query
   *
   * @type {string}
   * @memberof IGraph
   */
  readonly version: string;

  /**
   * returns a new instance of the Graph using the same
   * client within the context of the provider.
   *
   * @param {Element} component
   * @returns {IGraph}
   * @memberof IGraph
   */
  forComponent(component: Element): IGraph;

  /**
   * use this method to make calls directly to the Graph.
   *
   * @param {string} path
   * @returns {GraphRequest}
   * @memberof IGraph
   */
  api(path: string): GraphRequest;

  ///
  /// BATCH
  ///

  /**
   * creates a new batch request
   *
   * @returns {Batch}
   * @memberof IGraph
   */
  createBatch(): Batch;

  ///
  /// USER
  ///

  /**
   * async promise, returns Graph User data relating to the user logged in
   *
   * @returns {Promise<User>}
   * @memberof IGraph
   */
  getMe(): Promise<User>;

  /**
   * async promise, returns all Graph users associated with the userPrincipleName provided
   *
   * @param {string} userPrincipleName
   * @returns {Promise<User>}
   * @memberof IGraph
   */
  getUser(userPrincipleName: string): Promise<User>;

  ///
  /// PHOTO
  ///

  /**
   * async promise, returns Graph photos associated with contacts of the logged in user
   * @param contactId
   */
  getContactPhoto(contactId: string): Promise<string>;

  /**
   * async promise, returns Graph photo associated with provided userId
   * @param userId
   */
  getUserPhoto(userId: string): Promise<string>;

  /**
   * async promise, returns Graph photo associated with the logged in user
   */
  myPhoto(): Promise<string>;

  ///
  /// PERSON
  ///

  /**
   * async promise, returns all Graph people who are most relevant contacts to the signed in user.
   *
   * @param {string} query
   * @returns {Promise<Person[]>}
   * @memberof IGraph
   */
  findPerson(query: string): Promise<Person[]>;

  /**
   * async promise to the Graph for People, by default, it will request the most frequent contacts for the signed in user.
   *
   * @returns {(Promise<Person[]>)}
   * @memberof IGraph
   */
  getPeople(): Promise<Person[]>;

  /**
   * async promise to the Graph for People, defined by a group id
   *
   * @param {string} groupId
   * @returns {(Promise<Person[]>)}
   * @memberof IGraph
   */
  getPeopleFromGroup(groupId: string): Promise<Person[]>;

  ///
  /// CONTACTS
  ///

  /**
   * async promise, returns a Graph contact associated with the email provided
   *
   * @param {string} email
   * @returns {(Promise<Contact[]>)}
   * @memberof IGraph
   */
  findContactByEmail(email: string): Promise<Contact[]>;

  /**
   * async promise, returns Graph contact and/or Person associated with the email provided
   * Uses: Graph.findPerson(email) and Graph.findContactByEmail(email)
   *
   * @param {string} email
   * @returns {(Promise<Array<Person | Contact>)}
   * @memberof IGraph
   */
  findUserByEmail(email: string): Promise<Array<Person | Contact>>;

  ///
  /// EVENTS
  ///

  /**
   * async promise, returns Calender events associated with either the logged in user or a specific groupId
   *
   * @param {Date} startDateTime
   * @param {Date} endDateTime
   * @param {string} [groupId]
   * @returns {Promise<Event[]}
   * @memberof IGraph
   */
  getEvents(startDateTime: Date, endDateTime: Date, groupId?: string): Promise<Event[]>;

  ///
  /// PLANNER
  ///

  /**
   * async promise, allows developer to create new Planner task
   *
   * @param {PlannerTask} newTask
   * @returns {Promise<any>}
   * @memberof IGraph
   */
  addPlannerTask(newTask: PlannerTask): Promise<any>;

  /**
   * async promise, allows developer to assign people to task
   *
   * @param {string} taskId
   * @param {*} people
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof IGraph
   */
  assignPeopleToPlannerTask(taskId: string, people: any, eTag: string): Promise<any>;

  /**
   * async promise, returns all planner plans associated with the user logged in
   *
   * @returns {Promise<PlannerPlan[]>}
   * @memberof IGraph
   */
  getAllMyPlannerPlans(): Promise<PlannerPlan[]>;

  /**
   * async promise, returns bucket (for tasks) associated with a planId
   *
   * @param {string} planId
   * @returns {Promise<PlannerBucket[]>}
   * @memberof IGraph
   */
  getBucketsForPlannerPlan(planId: string): Promise<PlannerBucket[]>;

  /**
   * async promise, returns all planner plans associated with the group id
   *
   * @param {string} groupId
   * @returns {(Promise<PlannerPlan[]>)}
   * @memberof IGraph
   */
  getPlansForGroup(groupId: string): Promise<PlannerPlan[]>;

  /**
   * async promise, returns a single plan from the Graph associated with the planId
   *
   * @param {string} planId
   * @returns {Promise<PlannerPlan>}
   * @memberof BaseGraph
   */
  getSinglePlannerPlan(planId: string): Promise<PlannerPlan>;

  /**
   * async promise, returns all tasks from planner associated with a bucketId
   *
   * @param {string} bucketId
   * @returns {(Promise<PlannerTask[][]>)}
   * @memberof IGraph
   */
  getTasksForPlannerBucket(bucketId: string): Promise<PlannerTask[][]>;

  /**
   * async promise, allows developer to remove Planner task associated with taskId
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof IGraph
   */
  removePlannerTask(taskId: string, eTag: string): Promise<any>;

  /**
   * async promise, allows developer to set a task to complete, associated with taskId
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof IGraph
   */
  setPlannerTaskComplete(taskId: string, eTag: string): Promise<any>;

  /**
   * async promise, allows developer to set a task to incomplete, associated with taskId
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof IGraph
   */
  setPlannerTaskIncomplete(taskId: string, eTag: string): Promise<any>;

  /**
   * async promise, allows developer to set details of planner task associated with a taskId
   *
   * @param {string} taskId
   * @param {(PlannerTask)} details
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof IGraph
   */
  setPlannerTaskDetails(taskId: string, details: PlannerTask, eTag: string): Promise<any>;
}
