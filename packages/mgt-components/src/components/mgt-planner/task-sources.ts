/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/* eslint-disable max-classes-per-file */

import { IGraph, BetaGraph } from '@microsoft/mgt-element';
import { PlannerAssignments, PlannerBucket, PlannerPlan, PlannerTask } from '@microsoft/microsoft-graph-types';
import {
  addPlannerTask,
  assignPeopleToPlannerTask,
  getAllMyPlannerPlans,
  getBucketsForPlannerPlan,
  getPlansForGroup,
  getSinglePlannerPlan,
  getTasksForPlannerBucket,
  removePlannerTask,
  setPlannerTaskComplete,
  setPlannerTaskIncomplete
} from './graph.planner';

/**
 * ITask
 *
 * @export
 * @interface ITask
 */
export interface ITask {
  /**
   * id
   *
   * @type {string}
   * @memberof ITask
   */
  id: string;
  /**
   * name
   *
   * @type {string}
   * @memberof ITask
   */
  name: string;
  /**
   * task dueDate
   *
   * @type {Date}
   * @memberof ITask
   */
  dueDate: Date;
  /**
   * is task completed
   *
   * @type {boolean}
   * @memberof ITask
   */
  completed: boolean;
  /**
   * task topParentId
   *
   * @type {string}
   * @memberof ITask
   */
  topParentId: string;
  /**
   * task's immediate parent task id
   *
   * @type {string}
   * @memberof ITask
   */
  immediateParentId: string;
  /**
   * assignments
   *
   * @type {PlannerAssignments}
   * @memberof ITask
   */
  assignments: PlannerAssignments;
  /**
   * eTag
   *
   * @type {string}
   * @memberof ITask
   */
  eTag: string;
  /**
   * raw
   *
   * @type {PlannerTask}
   * @memberof ITask
   */
  _raw?: PlannerTask;
}
/**
 * container for tasks
 *
 * @export
 * @interface ITaskFolder
 */
export interface ITaskFolder {
  /**
   * id
   *
   * @type {string}
   * @memberof ITaskFolder
   */
  id: string;
  /**
   * name
   *
   * @type {string}
   * @memberof ITaskFolder
   */
  name: string;
  /**
   * parentId
   *
   * @type {string}
   * @memberof ITaskFolder
   */
  parentId: string;
  /**
   * raw
   *
   * @type {*}
   * @memberof ITaskFolder
   */
  _raw?: PlannerBucket;
}

/**
 * container for folders
 *
 * @export
 * @interface ITaskGroup
 */
export interface ITaskGroup {
  /**
   * string
   *
   * @type {string}
   * @memberof ITaskGroup
   */
  id: string;
  /**
   * secondaryId
   *
   * @type {string}
   * @memberof ITaskGroup
   */
  secondaryId?: string;
  /**
   * title
   *
   * @type {string}
   * @memberof ITaskGroup
   */
  title: string;
  /**
   * raw
   *
   * @type {*}
   * @memberof ITaskGroup
   */
  _raw?: PlannerPlan;

  /**
   * Plan Container ID. Same as the group ID of the group in the plan.
   *
   * @type {string}
   * @memberof ITaskGroup
   */
  containerId?: string;
}
/**
 * A common interface for both planner and todo tasks
 *
 * @export
 * @interface ITaskSource
 */
export interface ITaskSource {
  /**
   * Promise that returns task collections for the signed in user
   *
   * @returns {Promise<ITaskGroup[]>}
   * @memberof ITaskSource
   */
  getTaskGroups(): Promise<ITaskGroup[]>;

  /**
   * Promise that returns task collections for group id
   *
   * @returns {Promise<ITaskGroup[]>}
   * @memberof ITaskSource
   */
  getTaskGroupsForGroup(id: string): Promise<ITaskGroup[]>;

  /**
   * Promise that returns a single task collection by collection id
   *
   * @param {string} id
   * @returns {Promise<ITaskGroup>}
   * @memberof ITaskSource
   */
  getTaskGroup(id: string): Promise<ITaskGroup>;

  /**
   * Promise that returns all task groups in task collection
   *
   * @param {string} id
   * @returns {Promise<ITaskFolder[]>}
   * @memberof ITaskSource
   */
  getTaskFoldersForTaskGroup(id: string): Promise<ITaskFolder[]>;

  /**
   * Promise that returns all tasks in task group
   *
   * @param {string} id
   * @param {string} parId
   * @returns {Promise<ITask[]>}
   * @memberof ITaskSource
   */
  getTasksForTaskFolder(id: string, parId: string): Promise<ITask[]>;

  /**
   * Promise that completes a single task
   *
   * @param {ITask} task
   * @returns {Promise<any>}
   * @memberof ITaskSource
   */
  setTaskComplete(task: ITask): Promise<void>;

  /**
   * Promise that sets a task to incomplete
   *
   * @param {ITask} task
   * @returns {Promise<any>}
   * @memberof ITaskSource
   */
  setTaskIncomplete(task: ITask): Promise<void>;

  /**
   * Promise to add a new task
   *
   * @param {ITask} newTask
   * @returns {Promise<any>}
   * @memberof ITaskSource
   */
  addTask(newTask: ITask): Promise<PlannerTask>;

  /**
   * assign id's to task
   *
   * @param {ITask} task
   * @param {PlannerAssignments} people
   * @returns {Promise<any>}
   * @memberof ITaskSource
   */
  assignPeopleToTask(task: ITask, people: PlannerAssignments): Promise<void>;

  /**
   * Promise to delete a task by id
   *
   * @param {ITask} task
   * @returns {Promise<any>}
   * @memberof ITaskSource
   */
  removeTask(task: ITask): Promise<void>;

  /**
   * assigns task to the current signed in user
   *
   * @param {ITask} task
   * @param {string} myId
   * @returns {Boolean}
   * @memberof ITaskSource
   */
  isAssignedToMe(task: ITask, myId: string): boolean;
}
/**
 * async method to get user details
 *
 * @class TaskSourceBase
 */
class TaskSourceBase {
  /**
   * the IGraph instance to use for making Graph requests
   *
   * @type {IGraph}
   * @memberof TaskSourceBase
   */
  public graph: IGraph;

  constructor(graph: IGraph) {
    // Use an instance of BetaGraph since we know we need to call beta apis.
    this.graph = BetaGraph.fromGraph(graph);
  }
}

/**
 * Create Planner
 *
 * @export
 * @class PlannerTaskSource
 * @extends {TaskSourceBase}
 * @implements {ITaskSource}
 */
export class PlannerTaskSource extends TaskSourceBase implements ITaskSource {
  /**
   * returns promise with all of users plans
   *
   * @returns {Promise<ITaskGroup[]>}
   * @memberof PlannerTaskSource
   */
  public async getTaskGroups(): Promise<ITaskGroup[]> {
    const plans = await getAllMyPlannerPlans(this.graph);
    return plans.map(
      plan => ({ id: plan.id, title: plan.title, containerId: plan?.container?.containerId }) as ITaskGroup
    );
  }

  /**
   * returns promise with all of plans for group id
   *
   * @param {string} id
   * @returns {Promise<ITaskGroup[]>}
   * @memberof PlannerTaskSource
   */
  public async getTaskGroupsForGroup(id: string): Promise<ITaskGroup[]> {
    const plans = await getPlansForGroup(this.graph, id);

    return plans.map(plan => ({ id: plan.id, title: plan.title }) as ITaskGroup);
  }

  /**
   * returns promise single TaskGroup or plan from plan.id
   *
   * @param {string} id
   * @returns {Promise<ITaskGroup>}
   * @memberof PlannerTaskSource
   */
  public async getTaskGroup(id: string): Promise<ITaskGroup> {
    const plan = await getSinglePlannerPlan(this.graph, id);

    return { id: plan.id, title: plan.title, _raw: plan };
  }

  /**
   * returns promise with Bucket for a plan from bucket.id
   *
   * @param {string} id
   * @returns {Promise<ITaskFolder[]>}
   * @memberof PlannerTaskSource
   */
  public async getTaskFoldersForTaskGroup(id: string): Promise<ITaskFolder[]> {
    const buckets = await getBucketsForPlannerPlan(this.graph, id);

    return buckets.map(
      bucket =>
        ({
          _raw: bucket,
          id: bucket.id,
          name: bucket.name,
          parentId: bucket.planId
        }) as ITaskFolder
    );
  }

  /**
   * get all task from a Bucket given task id
   *
   * @param {string} id
   * @returns {Promise<ITask[]>}
   * @memberof PlannerTaskSource
   */
  public async getTasksForTaskFolder(id: string): Promise<ITask[]> {
    const tasks = await getTasksForPlannerBucket(this.graph, id);

    return tasks.map(
      task =>
        ({
          _raw: task,
          assignments: task.assignments,
          completed: task.percentComplete === 100,
          dueDate: task.dueDateTime && new Date(task.dueDateTime),
          eTag: task['@odata.etag'] as string,
          id: task.id,
          immediateParentId: task.bucketId,
          name: task.title,
          topParentId: task.planId
        }) as ITask
    );
  }

  /**
   * set task in planner to complete state by id
   *
   * @param {ITask} task
   * @returns {Promise<any>}
   * @memberof PlannerTaskSource
   */
  public async setTaskComplete(task: ITask): Promise<void> {
    return await setPlannerTaskComplete(this.graph, task);
  }

  /**
   * set task in planner to incomplete state by id
   *
   * @param {ITask} task
   * @returns {Promise<any>}
   * @memberof PlannerTaskSource
   */
  public async setTaskIncomplete(task: ITask): Promise<void> {
    return setPlannerTaskIncomplete(this.graph, task);
  }

  /**
   * add new task to bucket
   *
   * @param {ITask} newTask
   * @returns {Promise<any>}
   * @memberof PlannerTaskSource
   */
  public async addTask(newTask: ITask): Promise<PlannerTask> {
    return await addPlannerTask(this.graph, {
      assignments: newTask.assignments,
      bucketId: newTask.immediateParentId,
      dueDateTime: newTask.dueDate?.toISOString(),
      planId: newTask.topParentId,
      title: newTask.name
    });
  }

  /**
   * Assigns people to task
   *
   * @param {ITask} task
   * @param {PlannerAssignments} people
   * @returns {Promise<any>}
   * @memberof PlannerTaskSource
   */
  public async assignPeopleToTask(task: ITask, people: PlannerAssignments): Promise<void> {
    return assignPeopleToPlannerTask(this.graph, task, people);
  }

  /**
   * remove task from bucket
   *
   * @param {ITask} task
   * @returns {Promise<any>}
   * @memberof PlannerTaskSource
   */
  public async removeTask(task: ITask): Promise<void> {
    return await removePlannerTask(this.graph, task);
  }

  /**
   * assigns task to the signed in user
   *
   * @param {ITask} task
   * @param {string} myId
   * @returns {boolean}
   * @memberof PlannerTaskSource
   */
  public isAssignedToMe(task: ITask, myId: string): boolean {
    const keys = Object.keys(task.assignments);
    return keys.includes(myId);
  }
}
