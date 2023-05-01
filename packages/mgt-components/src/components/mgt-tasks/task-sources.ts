/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph, BetaGraph } from '@microsoft/mgt-element';
import { PlannerAssignments } from '@microsoft/microsoft-graph-types';
import { OutlookTask, OutlookTaskFolder, OutlookTaskGroup } from '@microsoft/microsoft-graph-types-beta';
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
} from './mgt-tasks.graph.planner';
import {
  addTodoTask,
  getAllMyTodoGroups,
  getAllTodoTasksForFolder,
  getFoldersForTodoGroup,
  getSingleTodoGroup,
  removeTodoTask,
  setTodoTaskComplete,
  setTodoTaskIncomplete
} from './mgt-tasks.graph.todo';

/**
 * Itask
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
   * @type {*}
   * @memberof ITask
   */
  _raw?: any;
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
  _raw?: any;
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
  _raw?: any;

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
   * @param {string} id
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof ITaskSource
   */
  setTaskComplete(id: string, eTag: string): Promise<any>;

  /**
   * Promise that sets a task to incomplete
   *
   * @param {string} id
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof ITaskSource
   */
  setTaskIncomplete(id: string, eTag: string): Promise<any>;

  /**
   * Promise to add a new task
   *
   * @param {ITask} newTask
   * @returns {Promise<any>}
   * @memberof ITaskSource
   */
  addTask(newTask: ITask): Promise<any>;

  /**
   * assign id's to task
   *
   * @param {string} id
   * @param {*} people
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof ITaskSource
   */
  assignPeopleToTask(id: string, people: any, eTag: string): Promise<any>;

  /**
   * Promise to delete a task by id
   *
   * @param {string} id
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof ITaskSource
   */
  removeTask(id: string, eTag: string): Promise<any>;

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
// tslint:disable-next-line: max-classes-per-file
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
      plan => ({ id: plan.id, title: plan.title, containerId: plan?.container?.containerId } as ITaskGroup)
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

    return plans.map(plan => ({ id: plan.id, title: plan.title } as ITaskGroup));
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
        } as ITaskFolder)
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
          eTag: task['@odata.etag'],
          id: task.id,
          immediateParentId: task.bucketId,
          name: task.title,
          topParentId: task.planId
        } as ITask)
    );
  }

  /**
   * set task in planner to complete state by id
   *
   * @param {string} id
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof PlannerTaskSource
   */
  public async setTaskComplete(id: string, eTag: string): Promise<any> {
    return await setPlannerTaskComplete(this.graph, id, eTag);
  }

  /**
   * set task in planner to incomplete state by id
   *
   * @param {string} id
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof PlannerTaskSource
   */
  public async setTaskIncomplete(id: string, eTag: string): Promise<any> {
    return await setPlannerTaskIncomplete(this.graph, id, eTag);
  }

  /**
   * add new task to bucket
   *
   * @param {ITask} newTask
   * @returns {Promise<any>}
   * @memberof PlannerTaskSource
   */
  public async addTask(newTask: ITask): Promise<any> {
    return await addPlannerTask(this.graph, {
      assignments: newTask.assignments,
      bucketId: newTask.immediateParentId,
      dueDateTime: newTask.dueDate && newTask.dueDate.toISOString(),
      planId: newTask.topParentId,
      title: newTask.name
    });
  }

  /**
   * Assigns people to task
   *
   * @param {string} id
   * @param {string} eTag
   * @param {*} people
   * @returns {Promise<any>}
   * @memberof PlannerTaskSource
   */
  public async assignPeopleToTask(id: string, eTag: string, people: any): Promise<any> {
    return await assignPeopleToPlannerTask(this.graph, id, eTag, people);
  }

  /**
   * remove task from bucket
   *
   * @param {string} id
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof PlannerTaskSource
   */
  public async removeTask(id: string, eTag: string): Promise<any> {
    return await removePlannerTask(this.graph, id, eTag);
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

/**
 * determins outlook task group for data source
 *
 * @export
 * @class TodoTaskSource
 * @extends {TaskSourceBase}
 * @implements {ITaskSource}
 */
// tslint:disable-next-line: max-classes-per-file
export class TodoTaskSource extends TaskSourceBase implements ITaskSource {
  /**
   * get all Outlook task groups
   *
   * @returns {Promise<ITaskGroup[]>}
   * @memberof TodoTaskSource
   */
  public async getTaskGroups(): Promise<ITaskGroup[]> {
    const groups: OutlookTaskGroup[] = await getAllMyTodoGroups(this.graph);

    return groups.map(
      group =>
        ({
          _raw: group,
          id: group.id,
          secondaryId: group.groupKey,
          title: group.name
        } as ITaskGroup)
    );
  }
  /**
   * get a single OutlookTaskGroup from id
   *
   * @param {string} id
   * @returns {Promise<ITaskGroup>}
   * @memberof TodoTaskSource
   */
  public async getTaskGroup(id: string): Promise<ITaskGroup> {
    const group: OutlookTaskGroup = await getSingleTodoGroup(this.graph, id);

    return { id: group.id, secondaryId: group.groupKey, title: group.name, _raw: group };
  }
  /**
   * get all OutlookTaskFolder for group by id
   *
   * @param {string} id
   * @returns {Promise<ITaskFolder[]>}
   * @memberof TodoTaskSource
   */
  public async getTaskFoldersForTaskGroup(id: string): Promise<ITaskFolder[]> {
    const folders: OutlookTaskFolder[] = await getFoldersForTodoGroup(this.graph, id);

    return folders.map(
      folder =>
        ({
          _raw: folder,
          id: folder.id,
          name: folder.name,
          parentId: id
        } as ITaskFolder)
    );
  }
  /**
   * gets all tasks for OutLook Task Folder by id
   *
   * @param {string} id
   * @param {string} parId
   * @returns {Promise<ITask[]>}
   * @memberof TodoTaskSource
   */
  public async getTasksForTaskFolder(id: string, parId: string): Promise<ITask[]> {
    const tasks: OutlookTask[] = await getAllTodoTasksForFolder(this.graph, id);

    return tasks.map(
      task =>
        ({
          _raw: task,
          assignments: {},
          completed: !!task.completedDateTime,
          dueDate: task.dueDateTime && new Date(task.dueDateTime.dateTime + 'Z'),
          eTag: task['@odata.etag'],
          id: task.id,
          immediateParentId: id,
          name: task.subject,
          topParentId: parId
        } as ITask)
    );
  }

  /**
   * set task in planner to complete state by id
   *
   * @param {string} id
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof TodoTaskSource
   */
  public async setTaskComplete(id: string, eTag: string): Promise<any> {
    return await setTodoTaskComplete(this.graph, id, eTag);
  }

  /**
   * Assigns people to task
   *
   * @param {string} id
   * @param {string} eTag
   * @param {*} people
   * @returns {Promise<any>}
   * @memberof PlannerTaskSource
   */
  public async assignPeopleToTask(id: string, eTag: string, people: any): Promise<any> {
    return await assignPeopleToPlannerTask(this.graph, id, eTag, people);
  }
  /**
   * set task in planner to incomplete state by id
   *
   * @param {string} id
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof TodoTaskSource
   */
  public async setTaskIncomplete(id: string, eTag: string): Promise<any> {
    return await setTodoTaskIncomplete(this.graph, id, eTag);
  }
  /**
   * add new task to planner
   *
   * @param {ITask} newTask
   * @returns {Promise<any>}
   * @memberof TodoTaskSource
   */
  public async addTask(newTask: ITask): Promise<any> {
    const task = {
      parentFolderId: newTask.immediateParentId,
      subject: newTask.name
    } as OutlookTask;
    if (newTask.dueDate) {
      task.dueDateTime = {
        dateTime: newTask.dueDate.toISOString(),
        timeZone: 'UTC'
      };
    }
    return await addTodoTask(this.graph, task);
  }
  /**
   * remove task from planner by id
   *
   * @param {string} id
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof TodoTaskSource
   */
  public async removeTask(id: string, eTag: string): Promise<any> {
    return await removeTodoTask(this.graph, id, eTag);
  }

  /**
   * if task is assigned in to user logged in
   *
   * @param {ITask} task
   * @param {string} myId
   * @returns {boolean}
   * @memberof TodoTaskSource
   */
  public isAssignedToMe(task: ITask, myId: string): boolean {
    return true;
  }

  /**
   * returns promise with all of plans for group id
   *
   * @param {string} id
   * @returns {Promise<ITaskGroup[]>}
   * @memberof PlannerTaskSource
   */
  public async getTaskGroupsForGroup(id: string): Promise<ITaskGroup[]> {
    return undefined;
  }
}
