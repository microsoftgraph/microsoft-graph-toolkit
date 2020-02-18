import { GraphRequest } from '@microsoft/microsoft-graph-client';
import * as GraphTypes from '@microsoft/microsoft-graph-types';
import * as BetaTypes from '@microsoft/microsoft-graph-types-beta';
import { Batch } from '.';

/**
 * The common functions of the Graph
 *
 * @export
 * @interface IGraph
 */
export interface IGraph {
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
   * @returns {(Promise<GraphTypes.User | BetaTypes.User>)}
   * @memberof IGraph
   */
  getMe(): Promise<GraphTypes.User | BetaTypes.User>;

  /**
   * async promise, returns all Graph users associated with the userPrincipleName provided
   *
   * @param {string} userPrincipleName
   * @returns {(Promise<GraphTypes.User | BetaTypes.User>)}
   * @memberof IGraph
   */
  getUser(userPrincipleName: string): Promise<GraphTypes.User | BetaTypes.User>;

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
   * @returns {(Promise<GraphTypes.Person[] | BetaTypes.Person[]>)}
   * @memberof IGraph
   */
  findPerson(query: string): Promise<GraphTypes.Person[] | BetaTypes.Person[]>;

  /**
   * async promise to the Graph for People, by default, it will request the most frequent contacts for the signed in user.
   *
   * @returns {(Promise<GraphTypes.Person[] | BetaTypes.Person[]>)}
   * @memberof IGraph
   */
  getPeople(): Promise<GraphTypes.Person[] | BetaTypes.Person[]>;

  /**
   * async promise to the Graph for People, defined by a group id
   *
   * @param {string} groupId
   * @returns {(Promise<GraphTypes.Person[] | BetaTypes.Person[]>)}
   * @memberof IGraph
   */
  getPeopleFromGroup(groupId: string): Promise<GraphTypes.Person[] | BetaTypes.Person[]>;

  ///
  /// CONTACTS
  ///

  /**
   * async promise, returns a Graph contact associated with the email provided
   *
   * @param {string} email
   * @returns {(Promise<GraphTypes.Contact[] | BetaTypes.Contact[]>)}
   * @memberof IGraph
   */
  findContactByEmail(email: string): Promise<GraphTypes.Contact[] | BetaTypes.Contact[]>;

  /**
   * async promise, returns Graph contact and/or Person associated with the email provided
   * Uses: Graph.findPerson(email) and Graph.findContactByEmail(email)
   *
   * @param {string} email
   * @returns {(Promise<Array<GraphTypes.Person | GraphTypes.Contact> | Array<BetaTypes.Person | BetaTypes.Contact>>)}
   * @memberof IGraph
   */
  findUserByEmail(
    email: string
  ): Promise<Array<GraphTypes.Person | GraphTypes.Contact> | Array<BetaTypes.Person | BetaTypes.Contact>>;

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
   * @memberof IGraph
   */
  getEvents(startDateTime: Date, endDateTime: Date, groupId?: string): Promise<GraphTypes.Event[] | BetaTypes.Event[]>;

  ///
  /// PLANNER
  ///

  /**
   * async promise, allows developer to create new Planner task
   *
   * @param {(GraphTypes.PlannerTask | BetaTypes.PlannerTask)} newTask
   * @returns {Promise<any>}
   * @memberof IGraph
   */
  addPlannerTask(newTask: GraphTypes.PlannerTask | BetaTypes.PlannerTask): Promise<any>;

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
   * @returns {(Promise<GraphTypes.PlannerPlan[] | BetaTypes.PlannerPlan[]>)}
   * @memberof IGraph
   */
  getAllMyPlannerPlans(): Promise<GraphTypes.PlannerPlan[] | BetaTypes.PlannerPlan[]>;

  /**
   * async promise, returns bucket (for tasks) associated with a planId
   *
   * @param {string} planId
   * @returns {(Promise<GraphTypes.PlannerBucket[] | BetaTypes.PlannerBucket[]>)}
   * @memberof IGraph
   */
  getBucketsForPlannerPlan(planId: string): Promise<GraphTypes.PlannerBucket[] | BetaTypes.PlannerBucket[]>;

  /**
   * async promise, returns all planner plans associated with the group id
   *
   * @param {string} groupId
   * @returns {(Promise<GraphTypes.PlannerPlan[] | BetaTypes.PlannerPlan[]>)}
   * @memberof IGraph
   */
  getPlansForGroup(groupId: string): Promise<GraphTypes.PlannerPlan[] | BetaTypes.PlannerPlan[]>;

  /**
   * async promise, returns a single plan from the Graph associated with the planId
   *
   * @param {string} planId
   * @returns {Promise<PlannerPlan>}
   * @memberof BaseGraph
   */
  getSinglePlannerPlan(planId: string): Promise<GraphTypes.PlannerPlan | BetaTypes.PlannerPlan>;

  /**
   * async promise, returns all tasks from planner associated with a bucketId
   *
   * @param {string} bucketId
   * @returns {(Promise<GraphTypes.PlannerTask[] | BetaTypes.PlannerTask[]>)}
   * @memberof IGraph
   */
  getTasksForPlannerBucket(bucketId: string): Promise<GraphTypes.PlannerTask[] | BetaTypes.PlannerTask[]>;

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
   * @param {(GraphTypes.PlannerTask | BetaTypes.PlannerTask)} details
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof IGraph
   */
  setPlannerTaskDetails(
    taskId: string,
    details: GraphTypes.PlannerTask | BetaTypes.PlannerTask,
    eTag: string
  ): Promise<any>;

  ///
  /// TO-DO
  ///

  /**
   * async promise, allows developer to add new to-do task
   *
   * @param {*} newTask
   * @returns {Promise<BetaTypes.OutlookTask>}
   * @memberof IGraph
   */
  addTodoTask(newTask: any): Promise<BetaTypes.OutlookTask>;

  /**
   * async promise, returns all Outlook taskGroups associated with the logged in user
   *
   * @returns {Promise<BetaTypes.OutlookTaskGroup[]>}
   * @memberof IGraph
   */
  getAllMyTodoGroups(): Promise<BetaTypes.OutlookTaskGroup[]>;

  /**
   * async promise, returns all Outlook tasks associated with a taskFolder with folderId
   *
   * @param {string} folderId
   * @returns {Promise<BetaTypes.OutlookTask[]>}
   * @memberof IGraph
   */
  getAllTodoTasksForFolder(folderId: string): Promise<BetaTypes.OutlookTask[]>;

  /**
   * async promise, returns all Outlook taskFolders associated with groupId
   *
   * @param {string} groupId
   * @returns {Promise<BetaTypes.OutlookTaskFolder[]>}
   * @memberof IGraph
   */
  getFoldersForTodoGroup(groupId: string): Promise<BetaTypes.OutlookTaskFolder[]>;

  /**
   * async promise, returns to-do tasks from Outlook groups associated with a groupId
   *
   * @param {string} groupId
   * @returns {Promise<BetaTypes.OutlookTaskGroup>}
   * @memberof IGraph
   */
  getSingleTodoGroup(groupId: string): Promise<BetaTypes.OutlookTaskGroup>;

  /**
   * async promise, allows developer to remove task based on taskId
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof IGraph
   */
  removeTodoTask(taskId: string, eTag: string): Promise<any>;

  /**
   * async promise, allows developer to set to-do task to completed state
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<BetaTypes.OutlookTask>}
   * @memberof IGraph
   */
  setTodoTaskComplete(taskId: string, eTag: string): Promise<BetaTypes.OutlookTask>;

  /**
   * async promise, allows developer to set to-do task to incomplete state
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<BetaTypes.OutlookTask>}
   * @memberof IGraph
   */
  setTodoTaskIncomplete(taskId: string, eTag: string): Promise<BetaTypes.OutlookTask>;

  /**
   * async promise, allows developer to redefine to-do Task details associated with a taskId
   *
   * @param {string} taskId
   * @param {*} task
   * @param {string} eTag
   * @returns {Promise<BetaTypes.OutlookTask>}
   * @memberof IGraph
   */
  setTodoTaskDetails(taskId: string, task: any, eTag: string): Promise<BetaTypes.OutlookTask>;
}
