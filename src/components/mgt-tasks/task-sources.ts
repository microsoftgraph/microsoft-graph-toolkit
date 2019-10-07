/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { PlannerAssignments, User } from '@microsoft/microsoft-graph-types';
import { OutlookTask, OutlookTaskFolder, OutlookTaskGroup } from '@microsoft/microsoft-graph-types-beta';
import { Graph } from '../../Graph';
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
   * @type {string}
   * @memberof ITask
   */
  dueDate: string;
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
 * @interface IDrawer
 */
export interface IDrawer {
  /**
   * id
   *
   * @type {string}
   * @memberof IDrawer
   */
  id: string;
  /**
   * name
   *
   * @type {string}
   * @memberof IDrawer
   */
  name: string;
  /**
   * parentId
   *
   * @type {string}
   * @memberof IDrawer
   */
  parentId: string;
  /**
   * raw
   *
   * @type {*}
   * @memberof IDrawer
   */
  _raw?: any;
}

/**
 * container for drawers
 *
 * @export
 * @interface IDresser
 */
export interface IDresser {
  /**
   * string
   *
   * @type {string}
   * @memberof IDresser
   */
  id: string;
  /**
   * secondaryId
   *
   * @type {string}
   * @memberof IDresser
   */
  secondaryId?: string;
  /**
   * title
   *
   * @type {string}
   * @memberof IDresser
   */
  title: string;
  /**
   * raw
   *
   * @type {*}
   * @memberof IDresser
   */
  _raw?: any;
}
/**
 * ItaskSource
 *
 * @export
 * @interface ITaskSource
 */
export interface ITaskSource {
  // tslint:disable
  me(): Promise<User>;
  getMyDressers(): Promise<IDresser[]>;
  getSingleDresser(id: string): Promise<IDresser>;
  getDrawersForDresser(id: string): Promise<IDrawer[]>;
  getAllTasksForDrawer(id: string, parId: string): Promise<ITask[]>;

  setTaskComplete(id: string, eTag: string): Promise<any>;
  setTaskIncomplete(id: string, eTag: string): Promise<any>;

  addTask(newTask: ITask): Promise<any>;
  removeTask(id: string, eTag: string): Promise<any>;
  // tslint:enable
}
/**
 * async method to get user details
 *
 * @class TaskSourceBase
 */
class TaskSourceBase {
  constructor(public graph: Graph) {}
  /**
   * returns Promise with logged in user details
   *
   * @returns {Promise<User>}
   * @memberof TaskSourceBase
   */
  public async me(): Promise<User> {
    return await this.graph.getMe();
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
   * @returns {Promise<IDresser[]>}
   * @memberof PlannerTaskSource
   */
  public async getMyDressers(): Promise<IDresser[]> {
    const plans = await this.graph.planner_getAllMyPlans();

    return plans.map(plan => ({ id: plan.id, title: plan.title } as IDresser));
  }
  /**
   * returns promise single dresser or plan from plan.id
   *
   * @param {string} id
   * @returns {Promise<IDresser>}
   * @memberof PlannerTaskSource
   */
  public async getSingleDresser(id: string): Promise<IDresser> {
    const plan = await this.graph.planner_getSinglePlan(id);

    return { id: plan.id, title: plan.title, _raw: plan };
  }
  /**
   * returns promise with Bucket for a plan from bucket.id
   *
   * @param {string} id
   * @returns {Promise<IDrawer[]>}
   * @memberof PlannerTaskSource
   */
  public async getDrawersForDresser(id: string): Promise<IDrawer[]> {
    const buckets = await this.graph.planner_getBucketsForPlan(id);

    return buckets.map(
      bucket =>
        ({
          _raw: bucket,
          id: bucket.id,
          name: bucket.name,
          parentId: bucket.planId
        } as IDrawer)
    );
  }

  /**
   * get all task from a Bucket given task id
   *
   * @param {string} id
   * @returns {Promise<ITask[]>}
   * @memberof PlannerTaskSource
   */
  public async getAllTasksForDrawer(id: string): Promise<ITask[]> {
    const tasks = await this.graph.planner_getTasksForBucket(id);

    return tasks.map(
      task =>
        ({
          _raw: task,
          assignments: task.assignments,
          completed: task.percentComplete === 100,
          dueDate: task.dueDateTime,
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
    return await this.graph.planner_setTaskComplete(id, eTag);
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
    return await this.graph.planner_setTaskIncomplete(id, eTag);
  }

  /**
   * add new task to bucket
   *
   * @param {ITask} newTask
   * @returns {Promise<any>}
   * @memberof PlannerTaskSource
   */
  public async addTask(newTask: ITask): Promise<any> {
    return await this.graph.planner_addTask({
      assignments: newTask.assignments,
      bucketId: newTask.immediateParentId,
      dueDateTime: newTask.dueDate,
      planId: newTask.topParentId,
      title: newTask.name
    });
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
    return await this.graph.planner_removeTask(id, eTag);
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
   * @returns {Promise<IDresser[]>}
   * @memberof TodoTaskSource
   */
  public async getMyDressers(): Promise<IDresser[]> {
    const groups: OutlookTaskGroup[] = await this.graph.todo_getAllMyGroups();

    return groups.map(
      group =>
        ({
          _raw: group,
          id: group.id,
          secondaryId: group.groupKey,
          title: group.name
        } as IDresser)
    );
  }
  /**
   * get a single OutlookTaskGroup from id
   *
   * @param {string} id
   * @returns {Promise<IDresser>}
   * @memberof TodoTaskSource
   */
  public async getSingleDresser(id: string): Promise<IDresser> {
    const group: OutlookTaskGroup = await this.graph.todo_getSingleGroup(id);

    return { id: group.id, secondaryId: group.groupKey, title: group.name, _raw: group };
  }
  /**
   * get all OutlookTaskFolder for group by id
   *
   * @param {string} id
   * @returns {Promise<IDrawer[]>}
   * @memberof TodoTaskSource
   */
  public async getDrawersForDresser(id: string): Promise<IDrawer[]> {
    const folders: OutlookTaskFolder[] = await this.graph.todo_getFoldersForGroup(id);

    return folders.map(
      folder =>
        ({
          _raw: folder,
          id: folder.id,
          name: folder.name,
          parentId: id
        } as IDrawer)
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
  public async getAllTasksForDrawer(id: string, parId: string): Promise<ITask[]> {
    const tasks: OutlookTask[] = await this.graph.todo_getAllTasksForFolder(id);

    return tasks.map(
      task =>
        ({
          _raw: task,
          assignments: {},
          completed: !!task.completedDateTime,
          dueDate: task.dueDateTime && task.dueDateTime.dateTime,
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
    return await this.graph.todo_setTaskComplete(id, eTag);
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
    return await this.graph.todo_setTaskIncomplete(id, eTag);
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
      assignedTo: plannerAssignmentsToTodoAssign(newTask.assignments),
      parentFolderId: newTask.immediateParentId,
      subject: newTask.name
    } as OutlookTask;
    if (newTask.dueDate) {
      task.dueDateTime = { dateTime: newTask.dueDate, timeZone: 'UTC' };
    }
    return await this.graph.todo_addTask(task);
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
    return await this.graph.todo_removeTask(id, eTag);
  }
}
/**
 * Assignment method for user to task
 *
 * @param {PlannerAssignments} assignments
 * @returns {string}
 */
function plannerAssignmentsToTodoAssign(assignments: PlannerAssignments): string {
  return 'John Doe';
}
